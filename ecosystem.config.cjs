module.exports = {
  apps: [
    {
      name: "sp-mcp",
      script: "dist/index.js",
      cwd: "/opt/sp-mcp",
      env: {
        NODE_ENV: "production",
        TRANSPORT: "http",
        PORT: 3500,
      },
      env_file: "/opt/sp-mcp/.env",
      instances: 1,
      autorestart: true,
      watch: false,
      max_memory_restart: "256M",
      log_date_format: "YYYY-MM-DD HH:mm:ss",
      error_file: "/root/.pm2/logs/sp-mcp-error.log",
      out_file: "/root/.pm2/logs/sp-mcp-out.log",
      merge_logs: true,
    },
  ],
};
