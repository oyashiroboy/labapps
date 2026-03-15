require 'json'

class UnitToolController < ApplicationController
  DOCX_CONTENT_TYPES = ['application/vnd.openxmlformats-officedocument.wordprocessingml.document'].freeze

  PRESETS = {
    'jis' => {
      'name' => 'JIS（日本語論文）',
      'no_space_units' => ['%', '°', '′', '″', '℃'],
      'operator_space' => 'optional',
      'temperature' => 'celsius',
      'liter' => 'any'
    },
    'nature' => {
      'name' => 'Nature系',
      'no_space_units' => ['%', '°', '′', '″'],
      'operator_space' => 'required',
      'temperature' => 'degree_c',
      'liter' => 'upper'
    },
    'acs' => {
      'name' => 'ACS/MDPI',
      'no_space_units' => ['%', '°', '′', '″'],
      'operator_space' => 'required',
      'temperature' => 'degree_c',
      'liter' => 'upper'
    }
  }.freeze

  def index
    @step = session[:unit_tool_step] || 1
    @preset = session[:unit_tool_preset] || 'nature'
    @no_space_units = session[:unit_tool_no_space] || PRESETS[@preset]['no_space_units']
    @checks = session[:unit_tool_checks] || ['unit_space']
    @results = load_results_file
    @presets = PRESETS
  end

  def select_style
    preset = params[:preset].to_s
    unless PRESETS.key?(preset)
      flash[:alert] = '無効なスタイルです'
      return redirect_to(unit_tool_path)
    end

    session[:unit_tool_nonce] ||= SecureRandom.hex(16)
    session[:unit_tool_step] = 2
    session[:unit_tool_preset] = preset
    session[:unit_tool_no_space] = PRESETS[preset]['no_space_units'].dup

    redirect_to(unit_tool_path)
  end

  def confirm_units
    units_raw = params[:no_space_units].to_s.strip
    units = units_raw.split(/[,、\s]+/).map(&:strip).reject(&:blank?).uniq
    units = ['%'] if units.empty?

    session[:unit_tool_no_space] = units
    session[:unit_tool_step] = 3

    redirect_to(unit_tool_path)
  end

  def select_checks
    checks = []
    checks << 'unit_space'
    checks << 'operator_space' if params[:check_operator] == '1'
    checks << 'temperature' if params[:check_temperature] == '1'
    checks << 'liter' if params[:check_liter] == '1'
    checks << 'time_unit' if params[:check_time] == '1'
    checks << 'inequality' if params[:check_inequality] == '1'

    session[:unit_tool_checks] = checks
    session[:unit_tool_step] = 4

    redirect_to(unit_tool_path)
  end

  def check
    uploaded = params[:docx_file]
    error = validate_uploaded_file(
      uploaded: uploaded,
      allowed_extensions: %w[.docx],
      allowed_content_types: DOCX_CONTENT_TYPES,
      max_bytes: MAX_DOCX_UPLOAD_BYTES,
      label: 'Wordファイル'
    )
    if error
      flash[:alert] = error
      return redirect_to(unit_tool_path)
    end

    nonce = session[:unit_tool_nonce]
    unless nonce
      flash[:alert] = 'セッションが無効です。最初からやり直してください。'
      return redirect_to(unit_tool_path)
    end

    input_path = Rails.root.join('tmp', "unit_#{nonce}_input.docx").to_s
    File.binwrite(input_path, uploaded.read)

    preset = session[:unit_tool_preset] || 'nature'
    config = {
      'no_space_units' => session[:unit_tool_no_space] || PRESETS[preset]['no_space_units'],
      'operator_space' => PRESETS[preset]['operator_space'],
      'temperature' => PRESETS[preset]['temperature'],
      'liter' => PRESETS[preset]['liter'],
      'checks' => session[:unit_tool_checks] || ['unit_space']
    }

    script_path = Rails.root.join('python_scripts', 'unit_checker.py').to_s
    config_json = config.to_json
    stdout, stderr, status, timed_out = capture3_with_timeout('python3', script_path, input_path, config_json)

    if timed_out
      render_processing_failure('unit_tool timeout')
      return redirect_to(unit_tool_path)
    end

    unless status&.success?
      render_processing_failure("unit_tool python error: #{stderr.to_s.tr("\n", ' ')}")
      return redirect_to(unit_tool_path)
    end

    begin
      result = JSON.parse(stdout)
      Rails.logger.info("unit_checker result: issues=#{result['total_issues']}, errors=#{result.dig('summary', 'error')}, warnings=#{result.dig('summary', 'warning')}, infos=#{result.dig('summary', 'info')}")
    rescue JSON::ParserError
      flash[:alert] = '結果の解析に失敗しました'
      return redirect_to(unit_tool_path)
    end

    if result['error']
      flash[:alert] = result['error']
      return redirect_to(unit_tool_path)
    end

    save_results_file(result)

    session[:unit_tool_step] = 5
    flash[:notice] = "#{result['total_issues']}件の問題を検出しました"
    redirect_to(unit_tool_path)
  ensure
    File.delete(input_path) if input_path && File.exist?(input_path)
  end

  def reset
    reset_session_data
    flash[:notice] = 'リセットしました'
    redirect_to(unit_tool_path)
  end

  def export_csv
    results = load_results_file

    if results.nil? || results['issues'].empty?
      flash[:alert] = 'エクスポートする結果がありません'
      return redirect_to(unit_tool_path)
    end

    csv_lines = ["段落,種類,重要度,検出内容,修正候補,メッセージ"]
    results['issues'].each do |issue|
      para = issue['para_idx'].present? ? issue['para_idx'] + 1 : '-'
      csv_lines << [
        para,
        issue['rule'] || '',
        issue['severity'] || '',
        issue['original'] || '',
        issue['suggestion'] || '',
        issue['message'] || ''
      ].map { |v| "\"#{v.to_s.gsub('"', '""')}\"" }.join(',')
    end

    send_data csv_lines.join("\n"),
              filename: 'unit_check_results.csv',
              type: 'text/csv; charset=utf-8',
              disposition: 'attachment'
  end

  def back
    current_step = session[:unit_tool_step] || 1
    session[:unit_tool_step] = [1, current_step - 1].max
    redirect_to(unit_tool_path)
  end

  private

  TEMP_FILE_TTL = 10.minutes

  def results_file_path
    nonce = session[:unit_tool_nonce]
    return nil unless nonce

    Rails.root.join('tmp', "unit_#{nonce}_results.json").to_s
  end

  def save_results_file(data)
    path = results_file_path
    return unless path

    File.write(path, data.to_json)
  end

  def load_results_file
    path = results_file_path
    return nil unless path && File.exist?(path)

    if File.mtime(path) < TEMP_FILE_TTL.ago
      File.delete(path) rescue nil
      return nil
    end

    JSON.parse(File.read(path))
  rescue JSON::ParserError
    nil
  end

  def reset_session_data
    path = results_file_path
    File.delete(path) if path && File.exist?(path)

    session.delete(:unit_tool_step)
    session.delete(:unit_tool_nonce)
    session.delete(:unit_tool_preset)
    session.delete(:unit_tool_no_space)
    session.delete(:unit_tool_checks)
  end
end
