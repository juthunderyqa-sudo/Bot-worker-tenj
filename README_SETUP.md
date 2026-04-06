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

6. Optional schedule setting:
   - WORKING_DAYS=1,2,3,4,5,6
   - 1=Mon, 2=Tue, 3=Wed, 4=Thu, 5=Fri, 6=Sat, 7=Sun

7. New in this build:
   - only chosen working days are shown for booking
   - after admin approves or rejects a request, the user receives a status message
   - admin can remove an approved booking from the admin panel
