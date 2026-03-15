class VbaToolController < ApplicationController
  VBA_CONTENT_TYPES = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel.sheet.macroEnabled.12'
  ].freeze

  def index
    input_path = nil
    output_path = nil

    if request.post? && params[:file].present?
      uploaded_file = params[:file]
      error = validate_uploaded_file(
        uploaded: uploaded_file,
        allowed_extensions: %w[.xlsx .xlsm],
        allowed_content_types: VBA_CONTENT_TYPES,
        max_bytes: MAX_DOCX_UPLOAD_BYTES,
        label: 'Excelファイル'
      )

      if error
        flash[:alert] = error
        return redirect_to(vba_tool_path)
      end

      script_path = Rails.root.join('python_scripts', 'vba_rewriter.py').to_s
      unless File.exist?(script_path)
        flash[:alert] = '現在この機能は利用できません（処理スクリプト未配置）'
        Rails.logger.error("vba_rewriter.py not found: #{script_path}")
        return redirect_to(vba_tool_path)
      end

      input_path = safe_temp_path(prefix: 'vba_input', extension: '.upload')
      output_path = safe_temp_path(prefix: 'vba_output', extension: '.xlsx')
      safe_name = sanitize_filename_base("processed_#{uploaded_file.original_filename}", fallback: 'processed_file')
      download_filename = "#{safe_name}.xlsx"

      File.binwrite(input_path, uploaded_file.read)
      stdout, stderr, status, timed_out = capture3_with_timeout('python3', script_path, input_path, output_path)

      if timed_out
        render_processing_failure("vba_tool timeout: input=#{input_path}")
        return redirect_to(vba_tool_path)
      end

      unless status&.success?
        render_processing_failure("vba_tool python error: #{stderr.to_s.tr("\n", ' ')}")
        return redirect_to(vba_tool_path)
      end

      unless File.exist?(output_path)
        render_processing_failure("vba_tool missing output: stdout=#{stdout.to_s.tr("\n", ' ')}")
        return redirect_to(vba_tool_path)
      end

      send_data File.binread(output_path),
                filename: download_filename,
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                disposition: 'attachment'
    end
  ensure
    File.delete(input_path) if input_path && File.exist?(input_path)
    File.delete(output_path) if output_path && File.exist?(output_path)
  end

  def manual
  end

  def examples
  end
end