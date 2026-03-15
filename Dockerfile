# syntax=docker/dockerfile:1

# ============================================================
# Base Stage: 共通の依存関係
# ============================================================
FROM ruby:3.3-slim AS base

RUN apt-get update -qq && apt-get install -y \
    build-essential \
    libsqlite3-dev \
    libyaml-dev \
    libvips \
    nodejs \
    python3 \
    python3-pip \
    curl \
    git \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /myapp

# Python関係
COPY requirements.txt /myapp/requirements.txt
RUN pip3 install -r requirements.txt --break-system-packages

# ============================================================
# Development Stage: 開発環境
# ============================================================
FROM base AS development

ENV RAILS_ENV=development

# Gemfileをコピーしてbundle install（dev/test含む）
COPY Gemfile Gemfile.lock ./
RUN bundle install

# storageディレクトリ作成
RUN mkdir -p /myapp/storage tmp/pids

EXPOSE 3000

# 開発サーバー起動（自動リロード有効）
CMD ["bash", "-c", "rm -f tmp/pids/server.pid && bundle exec rails s -p 3000 -b '0.0.0.0'"]

# ============================================================
# Production Build Stage: 本番ビルド
# ============================================================
FROM base AS production-build

ENV RAILS_ENV=production

# Gemfileをコピーしてbundle install（本番用のみ）
COPY Gemfile Gemfile.lock ./
RUN bundle config set --local deployment 'true' && \
    bundle config set --local without 'development test' && \
    bundle install

# アプリケーションコードをコピー
COPY . .

# アセットプリコンパイル
RUN SECRET_KEY_BASE_DUMMY=1 bundle exec rails assets:precompile && \
    chmod +x bin/*

# storageディレクトリ作成
RUN mkdir -p /myapp/storage tmp/pids

# ============================================================
# Production Stage: 本番実行環境（軽量化）
# ============================================================
FROM ruby:3.3-slim AS production

RUN apt-get update -qq && apt-get install -y \
    libsqlite3-0 \
    libyaml-0-2 \
    libvips42 \
    python3 \
    python3-pip \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /myapp

# Python関係（本番用）
COPY requirements.txt /myapp/requirements.txt
RUN pip3 install -r requirements.txt --break-system-packages

# ビルド済みアプリをコピー
COPY --from=production-build /myapp /myapp
COPY --from=production-build /usr/local/bundle /usr/local/bundle

ENV RAILS_ENV=production \
    RAILS_SERVE_STATIC_FILES=true

EXPOSE 3000

# ヘルスチェック
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
  CMD curl -f http://localhost:3000/up || exit 1

# Thruster + Puma（本番用）
CMD ["./bin/thrust", "./bin/rails", "server", "-b", "0.0.0.0"]
