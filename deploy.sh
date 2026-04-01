#!/bin/bash
set -e

echo "=== SharePoint MCP Server — Hetzner Deploy ==="

cd /opt/sp-mcp

# 1. Install & build
echo "→ Installing dependencies..."
npm install --production=false
echo "→ Building..."
npm run build

# 2. Check .env exists
if [ ! -f .env ]; then
  echo "⚠  No .env file found. Copying template..."
  cp .env.example .env
  echo "⚠  EDIT .env with your Azure AD credentials before starting!"
  echo "   nano /opt/sp-mcp/.env"
  exit 1
fi

# 3. PM2
echo "→ Starting PM2 process..."
pm2 delete sp-mcp 2>/dev/null || true
pm2 start ecosystem.config.cjs
pm2 save

# 4. Nginx
echo "→ Installing Nginx config..."
cp nginx-sp-mcp.conf /etc/nginx/sites-available/sp-mcp.tutto.one
ln -sf /etc/nginx/sites-available/sp-mcp.tutto.one /etc/nginx/sites-enabled/
nginx -t && systemctl reload nginx

echo ""
echo "=== Done ==="
echo "1. Add DNS A record: sp-mcp.tutto.one → 95.216.39.52"
echo "2. Run: certbot --nginx -d sp-mcp.tutto.one"
echo "3. Test: curl https://sp-mcp.tutto.one/health"
echo "4. Add to Claude.ai as MCP connector: https://sp-mcp.tutto.one/mcp"
