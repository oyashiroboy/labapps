Rails.application.routes.draw do
  # --- トップページ (一覧) ---
  root 'home#index'

  # --- Excel VBA書き換えツール ---
  # /vba_tool でアクセスしたら indexアクションへ
  get '/vba_tool',          to: 'vba_tool#index'
  post '/vba_tool',         to: 'vba_tool#index'
  get '/vba_tool/manual',   to: 'vba_tool#manual'
  get '/vba_tool/examples', to: 'vba_tool#examples'

  # PPTツールのルート
  get 'ppt_tool', to: 'ppt_tool#index'
  post 'ppt_tool/analyze', to: 'ppt_tool#analyze'
  post 'ppt_tool/generate', to: 'ppt_tool#generate'
  get 'ppt_tool/export', to: 'ppt_tool#export'
  post 'ppt_tool/import', to: 'ppt_tool#import'

  # 略語チェッカーのルート
  get 'abbr_tool', to: 'abbr_tool#index'
  post 'abbr_tool/extract', to: 'abbr_tool#extract'
  post 'abbr_tool/search', to: 'abbr_tool#search'
  get 'abbr_tool/reset', to: 'abbr_tool#reset'
  get 'abbr_tool/export', to: 'abbr_tool#export_list'

  # 単位チェッカーのルート
  get 'unit_tool', to: 'unit_tool#index'
  post 'unit_tool/select_style', to: 'unit_tool#select_style'
  post 'unit_tool/confirm_units', to: 'unit_tool#confirm_units'
  post 'unit_tool/select_checks', to: 'unit_tool#select_checks'
  post 'unit_tool/check', to: 'unit_tool#check'
  get 'unit_tool/reset', to: 'unit_tool#reset'
  get 'unit_tool/export', to: 'unit_tool#export_csv'
  get 'unit_tool/back', to: 'unit_tool#back'

  # サーバーヘルスチェック用（デフォルトのまま）
  get "up" => "rails/health#show", as: :rails_health_check
end