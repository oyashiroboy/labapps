# config/initializers/cleanup_temp_files.rb
# アプリ起動時に古い一時ファイルを削除

Rails.application.config.after_initialize do
  next if Rails.env.test?

  Thread.new do
    sleep 5  # 起動完了を待つ
    
    tmp_dir = Rails.root.join('tmp')
    threshold = 10.minutes.ago
    deleted = 0

    # unit_tool, abbr_tool, ppt_tool の一時ファイルを削除
    %w[unit_* abbr_* ppt_*].each do |pattern|
      Dir.glob(tmp_dir.join(pattern)).each do |file|
        next unless File.file?(file)
        if File.mtime(file) < threshold
          File.delete(file) rescue nil
          deleted += 1
        end
      end
    end

    Rails.logger.info("[Cleanup] Deleted #{deleted} old temp files on startup") if deleted > 0
  end
end
