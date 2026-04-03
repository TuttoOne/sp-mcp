#!/bin/bash
set -e

# Configuration — set these for your environment
DOMAIN="${SP_MCP_DOMAIN:-sp-mcp.example.com}"
INSTALL_DIR="${SP_MCP_DIR:-/opt/sp-mcp}"

echo "=== SharePoint MCP Server — Deploy ==="
echo "Domain: $DOMAIN"
echo "Install dir: $INSTALL_DIR"

cd "$INSTALL_DIR"

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
  echo "   nano $INSTALL_DIR/.env"
  exit 1
fi

# 3. PM2
echo "→ Starting PM2 process..."
pm2 delete sp-mcp 2>/dev/null || true
pm2 start ecosystem.config.cjs
pm2 save

# 4. Nginx (optional — skip if not using reverse proxy)
if command -v nginx &>/dev/null; then
  echo "→ Installing Nginx config..."
  sed "s/sp-mcp.example.com/$DOMAIN/g" nginx-sp-mcp.conf > /etc/nginx/sites-available/"$DOMAIN"
  ln -sf /etc/nginx/sites-available/"$DOMAIN" /etc/nginx/sites-enabled/
  nginx -t && systemctl reload nginx
  echo ""
  echo "=== Done ==="
  echo "1. Add DNS A record: $DOMAIN → <your-server-ip>"
  echo "2. Run: certbot --nginx -d $DOMAIN"
  echo "3. Test: curl https://$DOMAIN/health"
  echo "4. Add to Claude.ai as MCP connector: https://$DOMAIN/mcp"
else
  echo ""
  echo "=== Done (no Nginx) ==="
  echo "Server running on http://localhost:${PORT:-3500}/mcp"
fi
