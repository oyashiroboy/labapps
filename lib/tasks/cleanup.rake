# lib/tasks/cleanup.rake
# 一時ファイルの定期クリーンアップ
#
# 使用方法:
#   rails cleanup:temp_files
#   
# cron設定例（10分おき）:
#   */10 * * * * cd /myapp && bundle exec rails cleanup:temp_files RAILS_ENV=production >> /myapp/log/cleanup.log 2>&1

namespace :cleanup do
  desc '10分以上経過した一時ファイルを削除'
  task temp_files: :environment do
    tmp_dir = Rails.root.join('tmp')
    deleted = 0
    threshold = 10.minutes.ago

    # unit_tool関連ファイル
    Dir.glob(tmp_dir.join('unit_*')).each do |file|
      if File.mtime(file) < threshold
        File.delete(file)
        deleted += 1
        puts "Deleted: #{File.basename(file)}"
      end
    end

    # abbr_tool関連ファイル
    Dir.glob(tmp_dir.join('abbr_*')).each do |file|
      if File.mtime(file) < threshold
        File.delete(file)
        deleted += 1
        puts "Deleted: #{File.basename(file)}"
      end
    end

    # ppt_tool関連ファイル
    Dir.glob(tmp_dir.join('ppt_*')).each do |file|
      if File.mtime(file) < threshold
        File.delete(file)
        deleted += 1
        puts "Deleted: #{File.basename(file)}"
      end
    end

    puts "Total deleted: #{deleted} files"
  end
end
