1. Replace your project files with this package.
2. Put GOOGLE_SERVICE_ACCOUNT_JSON and ACCESS_PASSWORD into Cloudflare secrets:
   - wrangler secret put GOOGLE_SERVICE_ACCOUNT_JSON
   - wrangler secret put ACCESS_PASSWORD
3. Deploy:
   - npm install
   - npx wrangler deploy
4. Open setup URL:
   - https://YOUR-WORKER.workers.dev/setup?key=techno-perspektyva-setup-2026
5. In Telegram, send /start.

Notes:
- Worker name is preserved as bot-worker-tenj.
- Existing KV namespace is preserved.
- Existing sheet ID, admin IDs, setup key, and webhook secret are preserved.
- Password is asked only once per registered user.
