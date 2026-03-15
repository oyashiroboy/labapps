require 'json'

class PptToolController < ApplicationController
  VALID_MODES = %w[analogous complementary square triad].freeze
  IMAGE_CONTENT_TYPES = %w[
    image/jpeg
    image/png
    image/gif
    image/webp
    image/bmp
    image/tiff
  ].freeze

  def export
    settings = session[:ppt_tool_settings]
    settings = {} unless settings.is_a?(Hash)

    data = JSON.pretty_generate(settings)
    send_data data,
              filename: 'ppt_tool_settings.json',
              type: 'application/json',
              disposition: 'attachment'
  end

  def import
    uploaded = params[:settings_file]
    error = validate_uploaded_file(
      uploaded: uploaded,
      allowed_extensions: %w[.json],
      allowed_content_types: ['application/json'],
      max_bytes: MAX_JSON_UPLOAD_BYTES,
      label: '設定JSONファイル'
    )
    if error
      flash[:alert] = error
      return redirect_to(ppt_tool_path)
    end

    begin
      parsed = JSON.parse(uploaded.read)
    rescue JSON::ParserError
      flash[:alert] = '設定JSONの形式が不正です'
      return redirect_to(ppt_tool_path)
    end

    unless parsed.is_a?(Hash)
      flash[:alert] = '設定JSONの形式が不正です（オブジェクトを期待）'
      return redirect_to(ppt_tool_path)
    end

    parsed_mode = normalized_mode(parsed['mode'])
    if parsed.key?('mode') && parsed['mode'].present? && parsed_mode.nil?
      flash[:alert] = 'modeは analogous/complementary/square/triad のいずれかで指定してください'
      return redirect_to(ppt_tool_path)
    end
    imported_mode = parsed_mode || 'analogous'

    palette = parsed['palette']
    if palette && (!palette.is_a?(Array) || palette.size != 6 || palette.any? { |c| c.to_s !~ /\A[0-9A-Fa-f]{6}\z/ })
      flash[:alert] = 'paletteは6個の16進RGB（例: "15A3DC"）で指定してください'
      return redirect_to(ppt_tool_path)
    end

    session[:ppt_tool_settings] = {
      'mode' => imported_mode,
      'palette' => palette&.map { |c| c.to_s.upcase },
      'filename_base' => sanitize_filename_base(parsed['filename_base'], fallback: 'Themed_Template'),
      'updated_at' => Time.current.iso8601
    }.compact

    flash[:notice] = '設定を読み込みました（画像なしでも生成できます）'
    redirect_to(ppt_tool_path)
  end

  def index
    @saved_settings = session[:ppt_tool_settings].is_a?(Hash) ? session[:ppt_tool_settings] : nil
    @saved_palette = @saved_settings&.fetch('palette', nil)
    @saved_mode = @saved_settings&.fetch('mode', nil)
    @saved_filename_base = @saved_settings&.fetch('filename_base', nil)

    @default_palette = %w[2563EB 3B82F6 60A5FA 1E40AF 1D4ED8 0284C7]
    @current_palette = @saved_palette || @default_palette
  end

  def analyze
    image = params[:image]
    mode = normalized_mode(params[:mode])

    unless mode
      flash[:alert] = '配色モードが不正です'
      return redirect_to(ppt_tool_path)
    end

    error = validate_uploaded_file(
      uploaded: image,
      allowed_extensions: %w[.png .jpg .jpeg .gif .webp .bmp .tif .tiff],
      allowed_content_types: IMAGE_CONTENT_TYPES,
      max_bytes: MAX_IMAGE_UPLOAD_BYTES,
      label: '画像ファイル'
    )
    if error
      flash[:alert] = error
      return redirect_to(ppt_tool_path)
    end

    input_path = safe_temp_path(prefix: 'ppt_input', extension: '.img')
    File.binwrite(input_path, image.read)

    script_path = Rails.root.join('python_scripts', 'ppt_themer.py').to_s
    stdout, stderr, status, timed_out = capture3_with_timeout('python3', script_path, input_path, '-', '-', mode)

    if timed_out
      render_processing_failure('ppt_tool analyze timeout')
      return redirect_to(ppt_tool_path)
    end

    unless status&.success?
      render_processing_failure("ppt_tool analyze error: #{stderr.to_s.tr("\n", ' ')}")
      return redirect_to(ppt_tool_path)
    end

    begin
      parsed = JSON.parse(stdout)
    rescue JSON::ParserError
      flash[:alert] = '解析結果の形式が不正です'
      return redirect_to(ppt_tool_path)
    end

    palette = parsed['palette']
    if !palette.is_a?(Array) || palette.size != 6 || palette.any? { |c| c.to_s !~ /\A[0-9A-Fa-f]{6}\z/ }
      flash[:alert] = '解析結果のpaletteが不正です'
      return redirect_to(ppt_tool_path)
    end

    session[:ppt_tool_settings] = {
      'mode' => mode,
      'palette' => palette.map { |c| c.to_s.upcase },
      'updated_at' => Time.current.iso8601
    }

    flash[:notice] = '画像から色を決定しました。必要なら微調整してダウンロードしてください。'
    redirect_to(ppt_tool_path)
  ensure
    File.delete(input_path) if input_path && File.exist?(input_path)
  end

  def generate
    saved_settings = session[:ppt_tool_settings].is_a?(Hash) ? session[:ppt_tool_settings] : {}
    saved_palette = saved_settings['palette']
    saved_mode = normalized_mode(saved_settings['mode']) || 'analogous'

    submitted_palette = Array(params[:palette]).map { |c| c.to_s.strip }
    submitted_palette = submitted_palette.select(&:present?)

    palette_hex = nil
    if submitted_palette.present?
      unless submitted_palette.size == 6 && submitted_palette.all? { |c| c.match?(/\A#?[0-9A-Fa-f]{6}\z/) }
        flash[:alert] = '色は6個の #RRGGBB 形式で指定してください'
        return redirect_to(ppt_tool_path)
      end
      palette_hex = submitted_palette.map { |c| c.delete_prefix('#').upcase }
    end

    palette_hex ||= saved_palette
    unless palette_hex.is_a?(Array) && palette_hex.size == 6
      flash[:alert] = '先に画像解析するか、6色を指定してください'
      return redirect_to(ppt_tool_path)
    end

    filename_base = params[:filename_base].to_s.strip
    filename_base = saved_settings['filename_base'].to_s.strip if filename_base.blank? && saved_settings['filename_base'].present?
    filename_base = sanitize_filename_base(filename_base, fallback: "Themed_#{Time.current.to_i}")
    output_filename = "#{filename_base}.potx"
    output_path = safe_temp_path(prefix: 'ppt_output', extension: '.potx')

    template_candidates = [
      Rails.root.join('python_scripts', 'template.pptx').to_s,
      Rails.root.join('python_scripts', 'テンプレート.pptx').to_s
    ]
    template_path = template_candidates.find { |path| File.exist?(path) }
    unless template_path
      flash[:alert] = "テンプレートPPTXが見つかりません: #{template_candidates.join(', ')}"
      Rails.logger.error(flash[:alert])
      return redirect_to(ppt_tool_path)
    end

    input_path = safe_temp_path(prefix: 'ppt_input_no_image', extension: '.bin')
    File.binwrite(input_path, '')

    script_path = Rails.root.join('python_scripts', 'ppt_themer.py').to_s
    palette_arg = palette_hex.join(',')
    stdout, stderr, status, timed_out = capture3_with_timeout(
      'python3',
      script_path,
      input_path,
      template_path,
      output_path,
      saved_mode,
      palette_arg
    )

    if timed_out
      render_processing_failure('ppt_tool generate timeout')
      return redirect_to(ppt_tool_path)
    end

    unless status&.success?
      render_processing_failure("ppt_tool generate error: #{stderr.to_s.tr("\n", ' ')}")
      return redirect_to(ppt_tool_path)
    end

    unless File.exist?(output_path) && File.size?(output_path)
      render_processing_failure("ppt_tool output missing: stdout=#{stdout.to_s.tr("\n", ' ')}")
      return redirect_to(ppt_tool_path)
    end

    begin
      parsed = JSON.parse(stdout)
      result_palette = parsed.is_a?(Hash) && parsed['palette'].is_a?(Array) ? parsed['palette'] : palette_hex
      session[:ppt_tool_settings] = {
        'mode' => saved_mode,
        'palette' => result_palette,
        'filename_base' => filename_base,
        'updated_at' => Time.current.iso8601
      }.compact
    rescue JSON::ParserError
      session[:ppt_tool_settings] = {
        'mode' => saved_mode,
        'palette' => palette_hex,
        'filename_base' => filename_base,
        'updated_at' => Time.current.iso8601
      }.compact
    end

    send_data File.binread(output_path),
              filename: output_filename,
              type: 'application/vnd.openxmlformats-officedocument.presentationml.template',
              disposition: 'attachment'
  ensure
    File.delete(input_path) if input_path && File.exist?(input_path)
    File.delete(output_path) if output_path && File.exist?(output_path)
  end

  private

  def normalized_mode(raw_mode)
    mode = raw_mode.to_s.strip
    return nil if mode.blank?
    return mode if VALID_MODES.include?(mode)

    nil
  end

end
