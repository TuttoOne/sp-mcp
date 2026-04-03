module.exports = {
  apps: [
    {
      name: "sp-mcp",
      script: "dist/index.js",
      cwd: __dirname,
      env: {
        NODE_ENV: "production",
        TRANSPORT: "http",
        PORT: 3500,
      },
      env_file: ".env",
      instances: 1,
      autorestart: true,
      watch: false,
      max_memory_restart: "256M",
      log_date_format: "YYYY-MM-DD HH:mm:ss",
      merge_logs: true,
    },
  ],
};
