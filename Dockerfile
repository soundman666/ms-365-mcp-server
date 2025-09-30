# Multi-stage build для оптимізації розміру образу
FROM node:20-alpine AS builder

# Встановлення build dependencies
RUN apk add --no-cache git python3 make g++

WORKDIR /app

# Копіювання package files
COPY package*.json ./

# Встановлення всіх залежностей (включно з dev)
RUN npm ci

# Копіювання вихідного коду
COPY . .

# Створення порожньої папки generated на випадок якщо команда не спрацює
RUN mkdir -p generated

# Генерація клієнтського коду (optional)
RUN npm run generate || echo "Generate command not found or failed, continuing..."

# Збірка проекту
RUN npm run build

# Production stage
FROM node:20-alpine AS production

# Встановлення runtime dependencies включно з libsecret для keytar
RUN apk add --no-cache \
    curl \
    ca-certificates \
    tzdata \
    libsecret \
    dbus \
    gnome-keyring

# Створення non-root користувача
RUN addgroup -g 1001 -S nodejs && \
    adduser -S nodejs -u 1001

WORKDIR /app

# Копіювання package.json для production залежностей
COPY package*.json ./

# Встановлення тільки production залежностей
RUN npm ci --only=production && npm cache clean --force

# Копіювання збудованого коду з builder stage
COPY --from=builder --chown=nodejs:nodejs /app/dist ./dist

# Копіювання generated папки (може бути порожньою)
COPY --from=builder --chown=nodejs:nodejs /app/generated ./generated

# Створення необхідних папок з правильними правами
RUN mkdir -p /app/logs /app/tmp /app/.cache && \
    chown -R nodejs:nodejs /app

# Встановлення власника всіх файлів
RUN chown -R nodejs:nodejs /app

# Переключення на non-root користувача
USER nodejs

# Відкриття порту
EXPOSE 3000

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=40s --retries=3 \
  CMD curl -f http://localhost:3000/auth/metadata || exit 1

# Запуск сервера з відключенням keytar для containerized environment
CMD ["node", "dist/index.js", "--org-mode", "--http", "3000"]