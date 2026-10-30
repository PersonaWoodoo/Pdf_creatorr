import os, asyncio, httpx
from pathlib import Path
for line in Path('.env').read_text().splitlines():
    if '=' in line and not line.startswith('#'):
        k,v=line.split('=',1);os.environ.setdefault(k,v)
async def main():
    token=os.environ['BOT_TOKEN'];url=os.environ['PUBLIC_URL'].rstrip('/')+'/telegram/webhook'
    secret=os.environ['WEBHOOK_SECRET']
    async with httpx.AsyncClient(timeout=20) as c:
        r=await c.post(f'https://api.telegram.org/bot{token}/setWebhook',json={'url':url,'secret_token':secret,'allowed_updates':['message','pre_checkout_query']});print(r.json())
asyncio.run(main())
