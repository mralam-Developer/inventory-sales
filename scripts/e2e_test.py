"""End-to-end test of webapp/index.html against the real Supabase project.
Needs an EMPTY database (the first account created becomes the owner). Creates stocktest.* users; delete them afterwards."""
import asyncio, random
from playwright.async_api import async_playwright
R = random.randint(100000, 999999)
OWNER = f"stocktest.owner{R}@example.com"; MEMBER = f"stocktest.member{R}@example.com"; PW = "Test-pass-123"
URL = "http://127.0.0.1:8765/index.html"
log = []
def ok(name, cond, extra=""):
    log.append((name, bool(cond))); print(("PASS " if cond else "FAIL ") + name, extra, flush=True)
async def signup(page, name, email):
    await page.goto(URL); await page.wait_for_selector("#af")
    await page.click("#as"); await page.fill("#an", name); await page.fill("#ae", email); await page.fill("#ap", PW); await page.click("#ag")
async def accept(d): await d.accept()

async def main():
    async with async_playwright() as p:
        b = await p.chromium.launch()
        phone = dict(viewport={"width": 430, "height": 932}, device_scale_factor=2, is_mobile=True, has_touch=True)
        c1 = await b.new_context(**phone); pg = await c1.new_page()
        errs = []; pg.on("pageerror", lambda e: errs.append(str(e)))
        await signup(pg, "Test Owner", OWNER); await pg.wait_for_selector("#stats", timeout=20000)
        ok("owner signup lands in app", True)
        ok("owner sees People tab and Add item", await pg.locator('[data-tab=people]').count() == 1 and await pg.locator('#add').count() == 1)
        for nm, cat, unit, op in [("Basmati Rice 25 kg", "Grains", "Bag", "120"), ("Sugar 50 kg", "Grains", "Bag", "80"), ("Cooking Oil 5 L <b>x</b>", "Oil & Ghee", "Carton", "60.5")]:
            await pg.click("#add"); await pg.fill("#in", nm); await pg.fill("#ic", cat); await pg.fill("#iu", unit); await pg.fill("#io", op)
            await pg.click("#is"); await pg.wait_for_selector(f'.item:has-text("{nm[:12]}")')
        ok("3 items listed, html escaped", await pg.locator(".item").count() == 3 and await pg.locator(".item b b").count() == 0)
        await pg.click("#add"); await pg.fill("#in", "sugar 50 KG "); await pg.click("#is"); await pg.wait_for_selector("#toast.on.bad")
        ok("duplicate name rejected", "already exists" in await pg.inner_text("#toast")); await pg.click("#ix")
        await pg.click('[data-tab=in]')
        await pg.select_option("#mi", label=await pg.locator('#mi option:has-text("Basmati")').inner_text())
        await pg.fill("#mq", "200"); await pg.fill("#mp", "Supplier A"); await pg.fill("#mr", "PI-001"); await pg.click("#ms"); await pg.wait_for_selector("#list .mv")
        ok("stock in listed", "+200" in await pg.inner_text("#list .mv .qty"))
        ok("supplier kept, quantity cleared", await pg.input_value("#mp") == "Supplier A" and await pg.input_value("#mq") == "")
        await pg.click('[data-tab=out]')
        await pg.select_option("#mi", label=await pg.locator('#mi option:has-text("Basmati")').inner_text())
        await pg.fill("#mq", "90.5"); await pg.fill("#mp", "Customer A"); await pg.click("#ms"); await pg.wait_for_selector("#list .mv")
        dialogs = []
        async def dismiss(d): dialogs.append(d.message); await d.dismiss()
        pg.once("dialog", dismiss)
        await pg.select_option("#mi", label=await pg.locator('#mi option:has-text("Sugar")').inner_text())
        await pg.fill("#mq", "500"); await pg.click("#ms"); await pg.wait_for_timeout(1500)
        ok("oversell asks first, cancel saves nothing", dialogs and "Only 80" in dialogs[0] and await pg.locator("#list .mv").count() == 1, str(dialogs))
        pg.once("dialog", accept); await pg.click("#ms"); await pg.wait_for_function("document.querySelectorAll('#list .mv').length==2")
        await pg.screenshot(path="/tmp/shots/m_stock_out.png", full_page=True)
        await pg.click('[data-tab=stock]'); await pg.wait_for_selector(".item")
        rice = await pg.locator('.item:has-text("Basmati") .qty').inner_text(); sugar = await pg.locator('.item:has-text("Sugar") .qty').inner_text()
        ok("rice 229.5, sugar -420 flagged", rice.split()[0] == "229.5" and sugar.split()[0] == "-420" and await pg.locator('.item:has-text("Sugar") .qty.neg').count() == 1, rice + sugar)
        ok("stats", (await pg.inner_text("#stats")).split() == "3 Items 2 In stock 0 Out of stock 1 Below zero".split())
        await pg.screenshot(path="/tmp/shots/m_stock.png")
        await pg.fill("#q", "oil"); ok("search filters", await pg.locator(".item").count() == 1); await pg.fill("#q", "")
        await pg.click('.item:has-text("Basmati")'); await pg.wait_for_selector("#ih .mv")
        ok("history shows 2 entries", await pg.locator("#ih .mv").count() == 2)
        pg.once("dialog", accept); await pg.click("#idl"); await pg.wait_for_selector("#dlg #toast.on.bad")
        ok("delete item with entries blocked, message shown inside dialog", "has stock entries" in await pg.inner_text("#toast"))
        await pg.screenshot(path="/tmp/shots/m_item_dialog.png")
        await pg.fill("#io", "100"); await pg.click("#is"); await pg.wait_for_function("document.querySelector('.item .qty').textContent.startsWith('209.5')")
        ok("edit opening stock recalculates", True)

        # ---- second person: view only ----
        c2 = await b.new_context(**phone); pm = await c2.new_page()
        await signup(pm, "Test Member", MEMBER); await pm.wait_for_selector("#pc")
        leak = await pm.evaluate("""async()=>{const a=await sb.from('items').select('*');const m=await sb.from('stock_movements').select('*');const v=await sb.from('item_stock').select('*');const p=await sb.from('profiles').select('*');
          const up=await sb.from('profiles').update({approved:true,role:'owner'}).eq('id',S.user.id).select();return {items:a.data?.length,moves:m.data?.length,view:v.data?.length,profiles:p.data?.length,upRows:up.data?.length||0}}""")
        ok("unapproved account sees nothing and cannot approve itself", leak == {"items": 0, "moves": 0, "view": 0, "profiles": 1, "upRows": 0}, str(leak))
        anon = await pm.evaluate("""async()=>{const c=supabase.createClient(SB_URL,SB_KEY,{auth:{persistSession:false,storageKey:'x'}});const a=await c.from('items').select('*');return a.data?.length??a.error.message}""")
        ok("anonymous visitor gets nothing", anon == 0 or isinstance(anon, str), str(anon))
        await pg.click('[data-tab=people]'); await pg.click("#rf"); await pg.wait_for_selector("[data-ap]")
        ok("pending badge shows 1", await pg.inner_text("#bd-people") == "1")
        await pg.screenshot(path="/tmp/shots/m_people.png")
        await pg.click("[data-ap]"); await pg.wait_for_selector('[data-ap][data-on="0"]')
        await pm.click("#pc"); await pm.wait_for_selector(".item")
        ok("approved member sees the 3 items", await pm.locator(".item").count() == 3)
        ok("member has no People tab and no Add item", await pm.locator('[data-tab=people]').count() == 0 and await pm.locator('#add').count() == 0)
        ok("member footer says view only", "(view only)" in await pm.inner_text(".foot"))
        await pm.click('[data-tab=out]'); await pm.wait_for_selector("#list .mv")
        ok("member sees entries but no form and no Delete", await pm.locator("#mf").count() == 0 and await pm.locator("[data-del]").count() == 0 and await pm.locator("#list .mv").count() == 2)
        await pm.screenshot(path="/tmp/shots/m_member_out.png")
        await pm.click('[data-tab=stock]'); await pm.click('.item:has-text("Basmati")'); await pm.wait_for_selector("#ih .mv")
        ok("member item view is read-only with history", await pm.locator("#dlg input").count() == 0 and await pm.locator("#ih .mv").count() == 2)
        await pm.screenshot(path="/tmp/shots/m_member_item.png"); await pm.click("#ix")
        w = await pm.evaluate("""async()=>{const it=S.items[0];const mv=S.moves[0];
          const a=await sb.from('items').insert({name:'hack'});const b2=await sb.from('stock_movements').insert({item_id:it.id,type:'out',quantity:1});
          const c=await sb.from('items').update({opening_stock:999}).eq('id',it.id).select();const d=await sb.from('stock_movements').delete().eq('id',mv.id).select();
          const e=await sb.from('items').delete().eq('id',S.items[2].id).select();const f=await sb.from('profiles').update({role:'owner'}).eq('id',S.user.id).select();
          return {insItem:!!a.error,insMove:!!b2.error,upd:c.data?.length||0,delMove:d.data?.length||0,delItem:e.data?.length||0,role:f.data?.length||0}}""")
        ok("member cannot write through the API either", w == {"insItem": True, "insMove": True, "upd": 0, "delMove": 0, "delItem": 0, "role": 0}, str(w))
        # owner deletes an entry; member refresh sees the change
        await pg.click('[data-tab=out]'); await pg.wait_for_selector("#list .mv")
        pg.once("dialog", accept); await pg.locator('#list .mv:has-text("Sugar") [data-del]').click(); await pg.wait_for_function("document.querySelectorAll('#list .mv').length==1")
        await pm.click("#rf"); await pm.wait_for_function("document.querySelector('#stats').innerText.replace(/\\s+/g,' ')=='3 Items 3 In stock 0 Out of stock 0 Below zero'")
        ok("owner delete works and member sees it after Refresh", True)

        # ---- PC layout ----
        await pg.set_viewport_size({"width": 1366, "height": 800}); await pg.wait_for_selector("#list table")
        ok("PC: stock out shows a table under the form", await pg.locator("#list tbody tr").count() == 1 and await pg.locator("#mf").count() == 1)
        box_l = await pg.locator("#list").bounding_box(); box_d = await pg.locator("#list [data-del]").bounding_box(); box_t = await pg.locator("#list table").bounding_box()
        ok("PC: table fits inside its card, Delete visible", box_t["width"] <= box_l["width"] + 1 and box_d["x"] + box_d["width"] <= box_l["x"] + box_l["width"], f"{box_t['width']} {box_l['width']}")
        await pg.set_viewport_size({"width": 920, "height": 800}); box_l = await pg.locator("#list").bounding_box(); box_t = await pg.locator("#list table").bounding_box()
        ok("PC narrow (920px): table still fits", box_t["width"] <= box_l["width"] + 1, f"{box_t['width']} {box_l['width']}"); await pg.screenshot(path="/tmp/shots/pc_narrow_out.png")
        await pg.set_viewport_size({"width": 1366, "height": 800})
        nav = await pg.locator("nav").bounding_box(); ok("PC: tabs at the top", nav["y"] < 120, str(nav))
        await pg.screenshot(path="/tmp/shots/pc_stock_out.png")
        await pg.click('[data-tab=in]'); await pg.wait_for_selector("#list table"); await pg.screenshot(path="/tmp/shots/pc_stock_in.png")
        await pg.click('[data-tab=stock]'); await pg.wait_for_selector("#list table")
        ok("PC: stock table has 3 rows, 7 columns", await pg.locator("#list tbody tr").count() == 3 and await pg.locator("#list thead th").count() == 7)
        await pg.screenshot(path="/tmp/shots/pc_stock.png")
        await pg.click('#list tbody tr >> nth=0'); await pg.wait_for_selector("#ih .mv"); await pg.screenshot(path="/tmp/shots/pc_item.png"); await pg.click("#ix")
        await pg.click('[data-tab=people]'); await pg.wait_for_selector("[data-ap]"); await pg.screenshot(path="/tmp/shots/pc_people.png")
        await pm.set_viewport_size({"width": 1366, "height": 800}); await pm.click('[data-tab=in]'); await pm.wait_for_selector("#list table")
        ok("PC member: table without Delete column", await pm.locator("#list thead th").count() == 6); await pm.screenshot(path="/tmp/shots/pc_member_in.png")
        # remove access, sign out/in
        pg.once("dialog", accept); await pg.click('[data-ap][data-on="0"]'); await pg.wait_for_selector('[data-ap][data-on="1"]')
        await pm.click("#rf"); await pm.wait_for_selector("#pc"); ok("removed member is locked out", True)
        await pg.click("#so"); await pg.wait_for_selector("#af"); await pg.fill("#ae", OWNER); await pg.fill("#ap", "wrong-pass"); await pg.click("#ag"); await pg.wait_for_selector("#toast.on.bad")
        ok("wrong password message", "Invalid" in await pg.inner_text("#toast"))
        await pg.fill("#ap", PW); await pg.click("#ag"); await pg.wait_for_selector("#stats"); ok("sign in works", True)
        ok("no script errors", not errs, str(errs))
        await b.close()
    print("\nTOTAL", len(log), "FAILED:", [n for n, c in log if not c])
asyncio.run(main())
