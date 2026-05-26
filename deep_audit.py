# -*- coding: utf-8 -*-
"""
SheetCodeCrest - Deep Dive Audit Pass 2
Captures modals, admin tabs, and responsive views.
Requires the dev server already running on :5173
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from playwright.sync_api import sync_playwright
import os, time

OUT = r"C:\Users\new\.gemini\antigravity-ide\brain\b9ee2349-1107-40ea-bae2-68f1f5a91591"
BASE = "http://localhost:5173"

def shot(page, name, full=True):
    path = os.path.join(OUT, f"audit_{name}.png")
    page.screenshot(path=path, full_page=full)
    print(f"  [OK] audit_{name}.png")
    return path

def warn(msg):
    print(f"  [SKIP] {msg}")

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(viewport={"width": 1440, "height": 900})
        page = ctx.new_page()

        errors_list = []
        warnings_list = []
        page.on("console", lambda m: errors_list.append(m.text) if m.type == "error" else (warnings_list.append(m.text) if m.type == "warning" else None))
        page.on("pageerror", lambda e: errors_list.append(str(e)))

        # ── 1. Landing ──────────────────────────────────────────────────────
        print("[1] Landing page")
        page.goto(BASE)
        page.wait_for_load_state("networkidle")
        time.sleep(1.5)
        shot(page, "01_landing_full")

        # ── 2. Dark mode ────────────────────────────────────────────────────
        print("[2] Dark mode toggle")
        # Find theme button by role or common IDs
        for sel in ["#theme-toggle","[data-theme-toggle]","button.theme-btn","button[title]"]:
            try:
                btns = page.locator(sel).all()
                if btns:
                    btns[0].click(timeout=1500)
                    time.sleep(0.4)
                    break
            except:
                pass
        shot(page, "02_landing_dark_mode")

        # ── 3. Scroll sections ──────────────────────────────────────────────
        print("[3] Page sections")
        page.evaluate("window.scrollTo(0, 500)")
        time.sleep(0.4); shot(page, "03_features_section", False)
        page.evaluate("window.scrollTo(0, 1100)")
        time.sleep(0.4); shot(page, "04_how_it_works", False)
        page.evaluate("window.scrollTo(0, 1800)")
        time.sleep(0.4); shot(page, "05_pricing_section", False)
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(0.4); shot(page, "06_footer", False)
        page.evaluate("window.scrollTo(0, 0)")

        # ── 4. Auth modal ───────────────────────────────────────────────────
        print("[4] Auth modal")
        time.sleep(0.3)
        opened = False
        for label in ["Get Started", "Sign Up", "Login", "Log In", "Start Free"]:
            try:
                page.get_by_text(label, exact=False).first.click(timeout=2000)
                opened = True
                break
            except:
                pass
        time.sleep(0.7)
        if opened:
            shot(page, "07_auth_modal_open", False)
            # Try switching to login tab
            for label in ["Log In", "Login", "Sign In"]:
                try:
                    page.get_by_text(label, exact=False).nth(1).click(timeout=1500)
                    time.sleep(0.4)
                    break
                except:
                    pass
            shot(page, "08_auth_login_tab", False)
            page.keyboard.press("Escape")
            time.sleep(0.4)

        # ── 5. Sign up with test account ────────────────────────────────────
        print("[5] Creating test user")
        for label in ["Get Started", "Sign Up", "Register"]:
            try:
                page.get_by_text(label, exact=False).first.click(timeout=2000)
                break
            except:
                pass
        time.sleep(0.5)
        # Switch to signup tab if present
        for label in ["Sign Up", "Register", "Create Account"]:
            try:
                page.get_by_text(label, exact=False).first.click(timeout=1500)
                time.sleep(0.3)
                break
            except:
                pass
        try:
            all_inputs = page.locator("input:visible").all()
            print(f"  Found {len(all_inputs)} visible inputs")
            vals = ["audituser","Audit User","9990009000","audit@test.com","Audit@1234"]
            for i, inp in enumerate(all_inputs[:5]):
                inp.clear()
                inp.fill(vals[i])
            shot(page, "09_signup_filled", False)
            page.locator("button[type='submit']:visible").first.click(timeout=3000)
            time.sleep(1.5)
        except Exception as e:
            warn(f"signup fill: {e}")
        shot(page, "10_post_auth_state", False)

        # ── 6. Main app logged-in ───────────────────────────────────────────
        print("[6] Main app state")
        shot(page, "11_main_app_full")

        # ── 7. Checkout modal ───────────────────────────────────────────────
        print("[7] Checkout / Pricing modal")
        checkout_opened = False
        for label in ["Upgrade to Pro", "Upgrade", "PRO", "Buy Now", "Subscribe"]:
            try:
                page.get_by_text(label, exact=False).first.click(timeout=2500)
                checkout_opened = True
                break
            except:
                pass
        time.sleep(0.8)
        if checkout_opened:
            shot(page, "12_checkout_modal", False)
            for tab in ["Razorpay", "Card", "UPI", "Net Banking"]:
                try:
                    page.get_by_text(tab, exact=False).first.click(timeout=1500)
                    time.sleep(0.3)
                    shot(page, f"13_checkout_{tab.lower().replace(' ','_')}", False)
                except:
                    pass
            page.keyboard.press("Escape")
            time.sleep(0.4)
        else:
            warn("Checkout button not found - taking fallback screenshot")
            shot(page, "12_checkout_not_found", False)

        # ── 8. Dashboard / Profile modal ────────────────────────────────────
        print("[8] User dashboard modal")
        for label in ["Dashboard", "Account", "Profile", "My Account"]:
            try:
                page.get_by_text(label, exact=False).first.click(timeout=2500)
                time.sleep(0.8)
                shot(page, "14_dashboard_modal", False)
                page.keyboard.press("Escape")
                time.sleep(0.3)
                break
            except:
                pass

        # ── 9. Admin console - all 6 tabs ───────────────────────────────────
        print("[9] Admin console")
        admin_opened = False
        for label in ["Admin", "Admin Console", "Admin Panel"]:
            try:
                page.get_by_text(label, exact=False).first.click(timeout=2500)
                admin_opened = True
                break
            except:
                pass
        time.sleep(0.8)
        if admin_opened:
            shot(page, "15_admin_users_tab", False)
            tabs = [
                ("Payments", "16_admin_payments"),
                ("Plans", "17_admin_plans"),
                ("Analytics", "18_admin_analytics"),
                ("Settings", "19_admin_settings"),
                ("Activity", "20_admin_activity"),
            ]
            for (tab_name, file_name) in tabs:
                try:
                    page.get_by_text(tab_name, exact=False).first.click(timeout=2000)
                    time.sleep(0.5)
                    shot(page, file_name, False)
                except Exception as e:
                    warn(f"Admin tab '{tab_name}': {e}")
            page.keyboard.press("Escape")
            time.sleep(0.3)
        else:
            warn("Admin button not found in current state (need admin credentials)")

        # ── 10. Mobile 390x844 ──────────────────────────────────────────────
        print("[10] Mobile 390x844")
        mob = browser.new_context(viewport={"width": 390, "height": 844})
        mp = mob.new_page()
        mp.goto(BASE); mp.wait_for_load_state("networkidle"); time.sleep(1)
        shot(mp, "21_mobile_landing")
        mp.evaluate("window.scrollTo(0,400)"); time.sleep(0.3)
        shot(mp, "22_mobile_scroll", False)
        # Open mobile auth
        for label in ["Get Started", "Sign Up", "Login"]:
            try:
                mp.get_by_text(label, exact=False).first.click(timeout=2000)
                time.sleep(0.6)
                shot(mp, "23_mobile_auth_modal", False)
                mp.keyboard.press("Escape")
                break
            except:
                pass
        mob.close()

        # ── 11. Tablet 768x1024 ─────────────────────────────────────────────
        print("[11] Tablet 768x1024")
        tab = browser.new_context(viewport={"width": 768, "height": 1024})
        tp = tab.new_page()
        tp.goto(BASE); tp.wait_for_load_state("networkidle"); time.sleep(1)
        shot(tp, "24_tablet_full")
        tab.close()

        # ── 12. Wide 2560x1440 ──────────────────────────────────────────────
        print("[12] Wide 2560x1440")
        wide = browser.new_context(viewport={"width": 2560, "height": 1440})
        wp = wide.new_page()
        wp.goto(BASE); wp.wait_for_load_state("networkidle"); time.sleep(1)
        shot(wp, "25_ultrawide", False)
        wide.close()

        browser.close()

        # ── Report ───────────────────────────────────────────────────────────
        print("\n" + "="*60)
        print("CONSOLE ERRORS FOUND:", len(errors_list))
        for e in errors_list[:20]:
            print(f"  ERR: {e[:120]}")
        print("\nCONSOLE WARNINGS FOUND:", len(warnings_list))
        for w in warnings_list[:10]:
            print(f"  WARN: {w[:100]}")
        print("\nAudit complete.")
        return errors_list, warnings_list

if __name__ == "__main__":
    run()
