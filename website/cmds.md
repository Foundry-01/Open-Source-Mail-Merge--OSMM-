# Commands for OSMM Website

## Development Commands
```bash
# Start development server
npm run dev

# Build for production
npm run build

# Preview production build locally
npm run preview
```

## Production Deployment Commands
```bash
# Navigate to project directory
cd /var/www/osmailmerge.com/website

# Install dependencies
npm install

# Build production version
npm run build

# Test nginx configuration
sudo nginx -t

# Restart nginx
sudo systemctl restart nginx

# Check nginx status
sudo systemctl status nginx

# Check if site is accessible
curl -I https://osmailmerge.com
```

## Nginx Commands
```bash
# Test nginx configuration
sudo nginx -t

# Reload nginx (without downtime)
sudo systemctl reload nginx

# Restart nginx
sudo systemctl restart nginx

# Check nginx status
sudo systemctl status nginx

# View nginx error logs
sudo tail -f /var/log/nginx/error.log

# View nginx access logs
sudo tail -f /var/log/nginx/access.log
```

## File Locations
- **Website Source**: `/var/www/osmailmerge.com/website/`
- **Production Build**: `/var/www/osmailmerge.com/website/dist/`
- **Nginx Config**: `/etc/nginx/sites-available/osmailmerge.com`
- **Nginx Enabled**: `/etc/nginx/sites-enabled/osmailmerge.com`

## Troubleshooting
```bash
# Check if Node.js is installed
node --version

# Check if npm is available
npm --version

# Fix permissions on node_modules if needed
chmod +x node_modules/.bin/*

# If build fails, clean and reinstall
rm -rf node_modules package-lock.json
npm install
```
