require 'json'

class AbbrToolController < ApplicationController
  DOCX_CONTENT_TYPES = ['application/vnd.openxmlformats-officedocument.wordprocessingml.document'].freeze
  ABBR_PATTERN = /\A[A-Za-z][A-Za-z0-9.-]{0,49}\z/
  MAX_FULL_NAME_LENGTH = 200
  MAX_SYNONYM_LENGTH = 80
  MAX_ABBREVIATIONS = 120
  MAX_ABBR_JSON_BYTES = 150_000

  def index
    @step = session[:abbr_tool_step] || 1
    @candidates = load_data_file('candidates') || []
    @selected_abbrs = load_data_file('selected') || []
    @search_results = load_data_file('results') || []
    @intro_para_idx = session[:abbr_tool_intro_idx]
    @total_paragraphs = session[:abbr_tool_total_paras]
  end

  def extract
    uploaded = params[:docx_file]
    max_len = (params[:max_len].presence || 20).to_i.clamp(5, 50)

    error = validate_uploaded_file(
      uploaded: uploaded,
      allowed_extensions: %w[.docx],
      allowed_content_types: DOCX_CONTENT_TYPES,
      max_bytes: MAX_DOCX_UPLOAD_BYTES,
      label: 'Wordファイル'
    )
    if error
      flash[:alert] = error
      return redirect_to(abbr_tool_path)
    end

    session[:abbr_tool_nonce] ||= SecureRandom.hex(16)
    input_path = input_docx_path
    File.binwrite(input_path, uploaded.read)

    script_path = Rails.root.join('python_scripts', 'abbr_checker.py').to_s
    stdout, stderr, status, timed_out = capture3_with_timeout('python3', script_path, 'extract', input_path, max_len.to_s)

    if timed_out
      render_processing_failure('abbr_tool extract timeout')
      cleanup_files
      return redirect_to(abbr_tool_path)
    end

    unless status&.success?
      render_processing_failure("abbr_tool extract error: #{stderr.to_s.tr("\n", ' ')}")
      cleanup_files
      return redirect_to(abbr_tool_path)
    end

    begin
      result = JSON.parse(stdout)
    rescue JSON::ParserError
      flash[:alert] = '抽出結果の解析に失敗しました'
      return redirect_to(abbr_tool_path)
    end

    if result['error']
      flash[:alert] = result['error']
      return redirect_to(abbr_tool_path)
    end

    save_data_file('candidates', result['candidates'])
    save_data_file('selected', [])
    save_data_file('results', [])

    session[:abbr_tool_step] = 2
    session[:abbr_tool_intro_idx] = result['intro_para_idx']
    session[:abbr_tool_total_paras] = result['total_paragraphs']

    flash[:notice] = "#{result['candidates'].size}件のカッコ内テキストを抽出しました"
    redirect_to(abbr_tool_path)
  end

  def search
    file_path = input_docx_path

    unless file_path && File.exist?(file_path)
      flash[:alert] = 'ファイルが見つかりません。最初からやり直してください。'
      reset_session_data
      return redirect_to(abbr_tool_path)
    end

    abbreviations = build_abbreviations(params[:abbreviations] || {})
    if abbreviations.empty?
      flash[:alert] = '少なくとも1つの有効な略語を選択してください'
      return redirect_to(abbr_tool_path)
    end

    if abbreviations.size > MAX_ABBREVIATIONS
      flash[:alert] = "略語の選択数が多すぎます（#{MAX_ABBREVIATIONS}件まで）"
      return redirect_to(abbr_tool_path)
    end

    abbr_json = abbreviations.to_json
    if abbr_json.bytesize > MAX_ABBR_JSON_BYTES
      flash[:alert] = '入力サイズが大きすぎます。略語数またはシノニム数を減らしてください'
      return redirect_to(abbr_tool_path)
    end

    script_path = Rails.root.join('python_scripts', 'abbr_checker.py').to_s
    stdout, stderr, status, timed_out = capture3_with_timeout('python3', script_path, 'search', file_path, abbr_json)

    if timed_out
      render_processing_failure('abbr_tool search timeout')
      return redirect_to(abbr_tool_path)
    end

    unless status&.success?
      render_processing_failure("abbr_tool search error: #{stderr.to_s.tr("\n", ' ')}")
      return redirect_to(abbr_tool_path)
    end

    begin
      result = JSON.parse(stdout)
    rescue JSON::ParserError
      flash[:alert] = '検索結果の解析に失敗しました'
      return redirect_to(abbr_tool_path)
    end

    if result['error']
      flash[:alert] = result['error']
      return redirect_to(abbr_tool_path)
    end

    save_data_file('selected', abbreviations)
    save_data_file('results', result['results'])

    session[:abbr_tool_step] = 3
    flash[:notice] = "#{abbreviations.size}件の略語を検索しました"
    redirect_to(abbr_tool_path)
  end

  def reset
    cleanup_files
    reset_session_data
    flash[:notice] = 'リセットしました'
    redirect_to(abbr_tool_path)
  end

  def export_list
    results = load_data_file('results') || []

    if results.empty?
      flash[:alert] = 'エクスポートする結果がありません'
      return redirect_to(abbr_tool_path)
    end

    lines = ["Abbreviations", ""]
    results.each do |r|
      next if r['full_name'].blank?
      lines << "#{r['abbr']}: #{r['full_name']}"
    end

    send_data lines.join("\n"),
              filename: 'abbreviations_list.txt',
              type: 'text/plain',
              disposition: 'attachment'
  end

  private

  def input_docx_path
    nonce = session[:abbr_tool_nonce]
    return nil unless nonce

    Rails.root.join('tmp', "abbr_#{nonce}_input.docx").to_s
  end

  def data_file_path(type)
    nonce = session[:abbr_tool_nonce]
    return nil unless nonce

    Rails.root.join('tmp', "abbr_#{nonce}_#{type}.json").to_s
  end

  def save_data_file(type, data)
    path = data_file_path(type)
    return unless path

    File.write(path, data.to_json)
  end

  def load_data_file(type)
    path = data_file_path(type)
    return nil unless path && File.exist?(path)

    JSON.parse(File.read(path))
  rescue JSON::ParserError
    nil
  end

  def build_abbreviations(abbr_params)
    abbreviations = []

    abbr_params.each_value do |data|
      next unless data[:selected] == '1'

      abbr = data[:abbr].to_s.strip
      next unless abbr.match?(ABBR_PATTERN)

      full_name = data[:full_name].to_s.strip
      full_name = '' if full_name.length > MAX_FULL_NAME_LENGTH

      synonyms = data[:synonyms].to_s.strip
      synonym_list = synonyms.split(/[,、]/).map(&:strip).reject(&:blank?)
      synonym_list = synonym_list.first(20).map { |item| item[0, MAX_SYNONYM_LENGTH] }

      abbreviations << {
        'abbr' => abbr,
        'full_name' => full_name,
        'synonyms' => synonym_list
      }
    end

    abbreviations
  end

  def cleanup_files
    nonce = session[:abbr_tool_nonce]
    return unless nonce

    %w[input.docx candidates.json selected.json results.json].each do |suffix|
      path = Rails.root.join('tmp', "abbr_#{nonce}_#{suffix}").to_s
      File.delete(path) if File.exist?(path)
    end
  end

  def reset_session_data
    cleanup_files
    session.delete(:abbr_tool_step)
    session.delete(:abbr_tool_nonce)
    session.delete(:abbr_tool_intro_idx)
    session.delete(:abbr_tool_total_paras)
  end
end
