# Survey to SharePoint

## Deployment Setup

Use this split deployment:

- GitHub Pages for the frontend
- Contabo VPS for the backend API

GitHub Pages cannot run `functions/api/submit.js`, so the backend must run on your server.

## Files Used

- `index.html`: frontend
- `config.js`: frontend API URL
- `functions/api/submit.js`: survey submission logic
- `server.mjs`: Node server for Contabo
- `package.json`: start command

## Step 1: Point the frontend to your Contabo API

Edit `config.js` and set your real API URL:

```js
window.SURVEY_CONFIG = {
  submitUrl: "https://api.yourdomain.com/api/submit"
};
```

If you do not have a separate API subdomain, you can also use your main domain, for example:

```js
window.SURVEY_CONFIG = {
  submitUrl: "https://yourdomain.com/api/submit"
};
```

## Step 2: Deploy the frontend to GitHub Pages

Push the repo to GitHub, then enable Pages.

In GitHub:

1. Open the repo.
2. Go to `Settings`.
3. Go to `Pages`.
4. Under `Source`, choose `Deploy from a branch`.
5. Select your branch, usually `main`.
6. Select the folder `/ (root)`.
7. Save.

After GitHub Pages finishes, your frontend will be live.

## Step 3: SSH into the Contabo server

Use your server IP and user:

```bash
ssh root@your-server-ip
```

Or use a non-root sudo user if that is how the VPS is set up.

## Step 4: Install Node.js, Nginx, and PM2 on Contabo

On Ubuntu:

```bash
apt update
apt install -y nginx curl
curl -fsSL https://deb.nodesource.com/setup_20.x | bash -
apt install -y nodejs
npm install -g pm2
```

## Step 5: Clone the repo on the server

Choose a deploy directory and clone the repo:

```bash
cd /var/www
git clone https://github.com/your-user/your-repo.git survey-to-sharepoint
cd survey-to-sharepoint
```

## Step 6: Create the backend environment file

Create a `.env` file in the repo root:

```bash
nano .env
```

Put your real values in it:

```env
PORT=3000
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
SITE_ID=your-site-id
DRIVE_ID=your-drive-id
TEMPLATE_ITEM_ID=your-template-item-id
RESPONSES_FOLDER_ID=your-responses-folder-id
```

## Step 7: Start the API with PM2

From the repo root:

```bash
set -a
. ./.env
set +a
pm2 start server.mjs --name survey-api
pm2 save
pm2 startup
```

Check that it is running:

```bash
pm2 status
pm2 logs survey-api
```

The app listens on port `3000` unless you change `PORT`.

## Step 8: Configure Nginx

Create an Nginx site config:

```bash
nano /etc/nginx/sites-available/survey-api
```

Use this config:

```nginx
server {
    listen 80;
    server_name api.yourdomain.com;

    location /api/submit {
        proxy_pass http://127.0.0.1:3000;
        proxy_http_version 1.1;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

Enable it:

```bash
ln -s /etc/nginx/sites-available/survey-api /etc/nginx/sites-enabled/survey-api
nginx -t
systemctl reload nginx
```

## Step 9: Add SSL

If your domain already points to the Contabo server, install HTTPS:

```bash
apt install -y certbot python3-certbot-nginx
certbot --nginx -d api.yourdomain.com
```

After that, keep `config.js` pointed to the `https://` URL.

## Step 10: Test the full flow

1. Open the GitHub Pages URL.
2. Fill out the survey.
3. Submit it.
4. Confirm the request reaches `https://api.yourdomain.com/api/submit`.
5. Confirm a new workbook appears in SharePoint.

## Updating Later

When you push frontend changes:

- GitHub Pages will redeploy the UI

When you push backend changes:

```bash
cd /var/www/survey-to-sharepoint
git pull
pm2 restart survey-api
```

## Important Note

The workbook template still has to match what the code expects:

- worksheet name: `Inputs_From_User`
- respondent range: `E1:E3`
- question range: `E5:G44`
