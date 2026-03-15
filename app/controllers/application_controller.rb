require 'open3'
require 'securerandom'
require 'timeout'

class ApplicationController < ActionController::Base
  MAX_DOCX_UPLOAD_BYTES = 25.megabytes
  MAX_IMAGE_UPLOAD_BYTES = 10.megabytes
  MAX_JSON_UPLOAD_BYTES = 1.megabyte
  PYTHON_TIMEOUT_SECONDS = 60

  # Only allow modern browsers supporting webp images, web push, badges, import maps, CSS nesting, and CSS :has.
  allow_browser versions: :modern

  # Changes to the importmap will invalidate the etag for HTML responses
  stale_when_importmap_changes

  private

  def validate_uploaded_file(uploaded:, allowed_extensions:, allowed_content_types:, max_bytes:, label:)
    return "#{label}を選択してください" unless uploaded.present? && uploaded.respond_to?(:read)

    filename = uploaded.original_filename.to_s
    extension = File.extname(filename).downcase
    unless allowed_extensions.include?(extension)
      return "#{label}の拡張子が不正です"
    end

    content_type = uploaded.content_type.to_s
    unless allowed_content_types.include?(content_type)
      return "#{label}のMIMEタイプが不正です"
    end

    size = uploaded.size.to_i
    return "#{label}のサイズが上限を超えています（#{max_bytes / 1.megabyte}MBまで）" if size <= 0 || size > max_bytes

    nil
  end

  def safe_temp_path(prefix:, extension:)
    nonce = SecureRandom.hex(12)
    Rails.root.join('tmp', "#{prefix}_#{Time.current.to_i}_#{nonce}#{extension}").to_s
  end

  def sanitize_filename_base(raw_name, fallback: 'download')
    base = raw_name.to_s.strip
    base = fallback if base.blank?
    base = base.sub(/\.[A-Za-z0-9]+\z/, '')
    base = base.gsub(/[^A-Za-z0-9._-]/, '_')
    base = base.gsub(/_+/, '_').gsub(/\A[._-]+|[._-]+\z/, '')
    base = fallback if base.blank?
    base[0, 80]
  end

  def capture3_with_timeout(*command_args, timeout_seconds: PYTHON_TIMEOUT_SECONDS)
    stdout = ''
    stderr = ''
    status = nil

    begin
      Timeout.timeout(timeout_seconds) do
        stdout, stderr, status = Open3.capture3(*command_args)
      end
      [stdout, stderr, status, false]
    rescue Timeout::Error
      [stdout, stderr, nil, true]
    end
  end

  def render_processing_failure(log_message)
    Rails.logger.error(log_message)
    flash[:alert] = '処理中にエラーが発生しました。入力内容を確認して再実行してください。'
  end
end
