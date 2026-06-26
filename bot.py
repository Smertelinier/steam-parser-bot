# -*- coding: utf-8 -*-
import asyncio
import json
import os
import random
from datetime import datetime, timedelta
from io import BytesIO
from urllib.parse import quote

from curl_cffi import requests
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import BufferedInputFile, LabeledPrice, PreCheckoutQuery
from bs4 import BeautifulSoup
from openpyxl import Workbook

TOKEN = os.getenv("BOT_TOKEN", "YOUR_BOT_TOKEN_HERE")
FREE_DAYS = 3
PRICE = "300⭐/мес"
ADMIN_ID = None

USER_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "users.json")
bot = Bot(token=TOKEN)
dp = Dispatcher()

_LAST_STEAM_REQUEST = 0.0

_users_cache = None
_users_dirty = False
_users_lock = asyncio.Lock()


def _load_users():
    if not os.path.exists(USER_FILE):
        return {}
    with open(USER_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def _save_users(users):
    import tempfile
    tmp = USER_FILE + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(users, f, ensure_ascii=False, indent=2)
        os.replace(tmp, USER_FILE)
    except PermissionError:
        pass


async def _flush_users():
    global _users_cache, _users_dirty
    async with _users_lock:
        if _users_dirty and _users_cache is not None:
            loop = asyncio.get_event_loop()
            await loop.run_in_executor(None, _save_users, dict(_users_cache))
            _users_dirty = False


async def _periodic_flush():
    while True:
        await asyncio.sleep(30)
        await _flush_users()


def _get_users():
    global _users_cache
    if _users_cache is None:
        _users_cache = _load_users()
    return _users_cache


def _mark_dirty():
    global _users_dirty
    _users_dirty = True


def get_or_create_user(user_id):
    users = _get_users()
    uid = str(user_id)
    if uid not in users:
        users[uid] = {
            "trial_start": datetime.now().isoformat(),
            "subscribed_until": None,
            "searches": 0,
            "last_seen": None,
        }
        _mark_dirty()
    else:
        defaults = {"searches": 0, "last_seen": None, "blocked": False}
        changed = False
        for k, v in defaults.items():
            if k not in users[uid]:
                users[uid][k] = v
                changed = True
        if changed:
            _mark_dirty()
    return users[uid]


def is_subscribed(user_id):
    user = get_or_create_user(user_id)
    if user.get("blocked"):
        return False
    if user.get("subscribed_until"):
        until = datetime.fromisoformat(user["subscribed_until"])
        if datetime.now() < until:
            return True
    trial_start = datetime.fromisoformat(user["trial_start"])
    if datetime.now() - trial_start < timedelta(days=FREE_DAYS):
        return True
    return False


def get_user_info(user_id):
    users = _get_users()
    uid = str(user_id)
    user = users.get(uid)
    if not user:
        return None
    trial_start = datetime.fromisoformat(user["trial_start"])
    now = datetime.now()
    trial_used = now - trial_start
    trial_left = timedelta(days=FREE_DAYS) - trial_used if trial_used < timedelta(days=FREE_DAYS) else timedelta(0)
    sub_until = user.get("subscribed_until")
    blocked = user.get("blocked", False)
    return {
        "trial_start": trial_start,
        "trial_left": trial_left,
        "trial_expired": trial_used >= timedelta(days=FREE_DAYS),
        "subscribed_until": datetime.fromisoformat(sub_until) if sub_until else None,
        "blocked": blocked,
        "active": is_subscribed(int(user_id)),
        "searches": user.get("searches", 0),
        "last_seen": user.get("last_seen"),
    }


_SCRAPER_SESSION = None

_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/121.0.0.0 Safari/537.36",
]


def _get_session():
    global _SCRAPER_SESSION
    if _SCRAPER_SESSION is None:
        _SCRAPER_SESSION = requests.Session()
    return _SCRAPER_SESSION


async def _rate_limited_request(url, max_retries=3):
    global _LAST_STEAM_REQUEST
    session = _get_session()
    for attempt in range(max_retries):
        now = datetime.now().timestamp()
        since_last = now - _LAST_STEAM_REQUEST
        if since_last < 2.0:
            await asyncio.sleep(2.0 - since_last)
        _LAST_STEAM_REQUEST = datetime.now().timestamp()
        headers = {"User-Agent": random.choice(_AGENTS)}
        loop = asyncio.get_event_loop()
        resp = await loop.run_in_executor(
            None, lambda: session.get(url, headers=headers, impersonate="chrome120", timeout=25)
        )
        if resp.status_code == 200:
            return resp
        if resp.status_code == 429:
            await asyncio.sleep(5 * (attempt + 1))
            continue
        return None
    return None


async def search_steam(query):
    url = f"https://store.steampowered.com/search/?term={quote(query)}&specials=1"
    resp = await _rate_limited_request(url)
    if not resp:
        return None
    soup = BeautifulSoup(resp.text, "html.parser")
    results = []
    for item in soup.select("[data-ds-appid]"):
        name_el = item.select_one(".title")
        price_el = item.select_one(".discount_final_price")
        discount_el = item.select_one(".discount_pct")
        link_el = item.select_one("a[href]")
        if name_el:
            name = name_el.text.strip()
            price = price_el.text.strip() if price_el else "N/A"
            discount = discount_el.text.strip() if discount_el else "-"
            link = link_el.get("href", "") if link_el else ""
            results.append({"name": name, "price": price, "discount": discount, "link": link})
    return results[:20] if results else None


def make_excel(games):
    wb = Workbook()
    ws = wb.active
    ws.title = "Steam Sales by SAI"
    ws.append(["Название", "Цена", "Скидка", "Ссылка"])
    for g in games:
        ws.append([g["name"], g["price"], g["discount"], g["link"]])
    ws.append([])
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    uid = message.from_user.id
    get_or_create_user(uid)
    await message.answer(
        f"🎮 Steam Sales Radar\n\n"
        f"📌 Отправь название игры — я найду скидки в Steam.\n"
        f"Пример: \"action\" или \"RPG\"\n\n"
        f"🎁 Первые {FREE_DAYS} дня — бесплатно.\n"
        f"💳 /pay — оплатить подписку\n"
        f"ℹ️ /help — помощь и контакты"
    )


@dp.message(Command("help"))
async def cmd_help(message: types.Message):
    await message.answer(
        f"ℹ️ Помощь\n\n"
        f"📤 Отправь любой текстовой запрос — бот найдёт игры со скидками в Steam.\n"
        f"Примеры: \"action\", \"RPG\", \"стратегия\"\n\n"
        f"📎 Пришлёт Excel с колонками:\n"
        f"Название, Цена, Скидка, Ссылка\n\n"
        f"🎁 {FREE_DAYS} дня — бесплатно\n"
        f"💳 {PRICE} — после триала\n"
        f"📩 Вопросы/оплата: @Saidikcs\n"
        f"🛠 Админ: @Saidikcs"
    )


@dp.message(Command("pay"))
async def cmd_pay(message: types.Message):
    keyboard = types.InlineKeyboardMarkup(
        inline_keyboard=[
            [types.InlineKeyboardButton(text="💳 Купить подписку 300⭐/мес", callback_data="buy_sub")],
        ]
    )
    await message.answer(
        f"💰 Подписка на Steam Sales Radar\n\n"
        f"⭐ 300 Stars в месяц\n"
        f"📦 Безлимитный поиск скидок\n"
        f"📊 Excel-отчёты\n\n"
        f"Нажми кнопку ниже для оплаты:",
        reply_markup=keyboard,
    )


@dp.callback_query(lambda c: c.data == "buy_sub")
async def buy_sub(callback: types.CallbackQuery):
    await callback.message.delete()
    await callback.message.answer_invoice(
        title="Steam Sales Radar — 1 месяц",
        description="Безлимитный поиск скидок и Excel-отчёты. 30 дней доступа.",
        payload=f"sub_{callback.from_user.id}",
        provider_token="",
        currency="XTR",
        prices=[LabeledPrice(label="Подписка 1 месяц", amount=300)],
    )
    await callback.answer()


@dp.pre_checkout_query()
async def pre_checkout(pre_checkout_q: PreCheckoutQuery):
    await pre_checkout_q.answer(ok=True)


@dp.message(F.successful_payment)
async def paid(message: types.Message):
    users = _get_users()
    uid = str(message.from_user.id)
    if uid not in users:
        users[uid] = {"trial_start": datetime.now().isoformat(), "subscribed_until": None}
    users[uid]["subscribed_until"] = (datetime.now() + timedelta(days=30)).isoformat()
    _mark_dirty()
    await _flush_users()
    await message.answer(
        f"✅ Оплата прошла! Подписка активирована на 30 дней.\n"
        f"Отправляй запросы — бот работает."
    )


@dp.message(Command("admin"))
async def cmd_admin(message: types.Message):
    if message.from_user.username != "Saidikcs":
        await message.answer("Нет доступа")
        return
    keyboard = types.InlineKeyboardMarkup(
        inline_keyboard=[
            [
                types.InlineKeyboardButton(text="Список пользователей", callback_data="admin_users"),
                types.InlineKeyboardButton(text="Продлить подписку", callback_data="admin_extend"),
            ],
            [types.InlineKeyboardButton(text="Статистика", callback_data="admin_stats")],
        ]
    )
    await message.answer("Админ-панель:", reply_markup=keyboard)


@dp.callback_query(lambda c: c.data == "admin_users")
async def admin_users(callback: types.CallbackQuery):
    if callback.from_user.username != "Saidikcs":
        return
    users = _get_users()
    if not users:
        await callback.message.edit_text("Нет пользователей")
        return
    lines = []
    for uid, data in users.items():
        trial = datetime.fromisoformat(data["trial_start"])
        active = "✅" if is_subscribed(int(uid)) else "❌"
        sub = data.get("subscribed_until", "-")
        lines.append(f"{active} ID: {uid}\n  Регистрация: {trial.strftime('%d.%m.%Y')}\n  Подписка до: {sub}")
    await callback.message.edit_text("Пользователи:\n\n" + "\n".join(lines))
    await callback.answer()


@dp.callback_query(lambda c: c.data == "admin_extend")
async def admin_extend(callback: types.CallbackQuery):
    if callback.from_user.username != "Saidikcs":
        return
    await callback.message.edit_text(
        "Чтобы продлить пользователю подписку на месяц, напиши:\n"
        "/extend USER_ID\n\n"
        "USER_ID можно узнать в списке пользователей."
    )
    await callback.answer()


@dp.callback_query(lambda c: c.data == "admin_stats")
async def admin_stats(callback: types.CallbackQuery):
    if callback.from_user.username != "Saidikcs":
        return
    users = _get_users()
    total = len(users)
    active = sum(1 for uid in users if is_subscribed(int(uid)))
    await callback.message.edit_text(f"Всего пользователей: {total}\nАктивных: {active}")
    await callback.answer()


@dp.message(Command("extend"))
async def cmd_extend(message: types.Message):
    if message.from_user.username != "Saidikcs":
        return
    parts = message.text.split()
    if len(parts) < 2:
        await message.answer("Использование: /extend USER_ID")
        return
    uid = parts[1]
    users = _get_users()
    if uid not in users:
        await message.answer("Пользователь не найден")
        return
    users[uid]["subscribed_until"] = (datetime.now() + timedelta(days=30)).isoformat()
    _mark_dirty()
    await _flush_users()
    await message.answer(f"Подписка продлена пользователю {uid} на 30 дней")


@dp.message(Command("mytrial"))
async def cmd_mytrial(message: types.Message):
    info = get_user_info(message.from_user.id)
    text = (
        f"📊 Статус:\n"
        f"🎁 Триал: осталось {info['trial_left'].days} дн {info['trial_left'].seconds//3600} ч\n"
        f"{'🔴 Истёк' if info['trial_expired'] else '🟢 Активен'}\n"
    )
    if info["subscribed_until"]:
        text += f"💳 Подписка до: {info['subscribed_until'].strftime('%d.%m.%Y')}\n"
    if info["blocked"]:
        text += "🚫 Заблокирован\n"
    text += f"\n💳 /pay — продлить подписку"
    await message.answer(text)


@dp.message(Command("stats"))
async def cmd_stats(message: types.Message):
    if message.from_user.username != "Saidikcs":
        return
    users = _get_users()
    now = datetime.now()
    total = len(users)
    active = 0
    on_trial = 0
    blocked = 0
    total_searches = 0
    for uid, u in users.items():
        total_searches += u.get("searches", 0)
        if u.get("blocked"):
            blocked += 1
            continue
        if u.get("subscribed_until"):
            until = datetime.fromisoformat(u["subscribed_until"])
            if now < until:
                active += 1
                continue
        trial_start = datetime.fromisoformat(u["trial_start"])
        if now - trial_start < timedelta(days=FREE_DAYS):
            on_trial += 1
            active += 1
    text = (
        f"📊 Статистика бота\n"
        f"👥 Всего пользователей: {total}\n"
        f"🟢 Активных: {active}\n"
        f"🎁 На триале: {on_trial}\n"
        f"💳 С подпиской: {active - on_trial}\n"
        f"🚫 Заблокировано: {blocked}\n"
        f"🔍 Всего поисков: {total_searches}"
    )
    await message.answer(text)


@dp.message(Command("check"))
async def cmd_check(message: types.Message):
    if message.from_user.username != "Saidikcs":
        return
    parts = message.text.split()
    if len(parts) < 2:
        await message.answer("Использование: /check USER_ID")
        return
    try:
        uid = int(parts[1]) if parts[1].isdigit() else parts[1]
        info = get_user_info(uid)
        if not info:
            await message.answer("Пользователь не найден")
            return
        text = (
            f"📊 Пользователь {uid}:\n"
            f"{'🟢 Активен' if info['active'] else '🔴 Неактивен'}\n"
            f"🎁 Триал начат: {info['trial_start'].strftime('%d.%m.%Y %H:%M')}\n"
            f"📅 Триал осталось: {info['trial_left'].days} дн\n"
        )
        if info["subscribed_until"]:
            text += f"💳 Подписка до: {info['subscribed_until'].strftime('%d.%m.%Y')}\n"
        if info["blocked"]:
            text += "🚫 Заблокирован\n"
        text += f"🔍 Поисков: {info['searches']}\n"
        if info["last_seen"]:
            text += f"🕐 Последний раз: {info['last_seen'][:16].replace('T', ' ')}"
        await message.answer(text)
    except Exception:
        await message.answer("Ошибка. ID должен быть числом.")


@dp.message(Command("block"))
async def cmd_block(message: types.Message):
    if message.from_user.username != "Saidikcs":
        return
    parts = message.text.split()
    if len(parts) < 2:
        await message.answer("Использование: /block USER_ID")
        return
    uid = parts[1]
    users = _get_users()
    if uid not in users:
        await message.answer("Пользователь не найден")
        return
    users[uid]["blocked"] = True
    _mark_dirty()
    await _flush_users()
    await message.answer(f"🚫 Пользователь {uid} заблокирован")


@dp.message(Command("unblock"))
async def cmd_unblock(message: types.Message):
    if message.from_user.username != "Saidikcs":
        return
    parts = message.text.split()
    if len(parts) < 2:
        await message.answer("Использование: /unblock USER_ID")
        return
    uid = parts[1]
    users = _get_users()
    if uid not in users:
        await message.answer("Пользователь не найден")
        return
    users[uid]["blocked"] = False
    _mark_dirty()
    await _flush_users()
    await message.answer(f"✅ Пользователь {uid} разблокирован")


@dp.message()
async def handle_search(message: types.Message):
    if not message.text or message.text.startswith("/"):
        return

    if not is_subscribed(message.from_user.id):
        await message.answer(
            f"😔 Бесплатный период ({FREE_DAYS} дня) закончился.\n"
            f"💳 /pay — продлить подписку за {PRICE}"
        )
        return

    await message.answer(f"🔍 Ищу скидки по запросу \"{message.text[:50]}\"...")

    users = _get_users()
    uid = str(message.from_user.id)
    if uid in users:
        users[uid]["searches"] = users[uid].get("searches", 0) + 1
        users[uid]["last_seen"] = datetime.now().isoformat()
        _mark_dirty()

    try:
        games = await search_steam(message.text)
    except Exception:
        await message.answer("Ошибка при поиске. Попробуй позже.")
        return
    if not games:
        await message.answer("Ничего не найдено. Попробуй другой запрос.")
        return

    excel = make_excel(games)
    top = games[0]
    caption = (
        f"✅ Найдено игр: {len(games)}\n\n"
        f"🏆 {top['name']} | {top['price']}\n\n"
        f"💳 /pay — подписка | ❓ @Saidikcs"
    )
    await message.answer_document(
        BufferedInputFile(excel.read(), filename=f"steam_sales.xlsx"),
        caption=caption,
    )


def migrate_all_users():
    users = _get_users()
    defaults = {"searches": 0, "last_seen": None, "blocked": False}
    changed = False
    for u in users.values():
        for k, v in defaults.items():
            if k not in u:
                u[k] = v
                changed = True
    if changed:
        _save_users(users)
        print("users.json обновлён: добавлены новые поля")


async def main():
    migrate_all_users()
    asyncio.create_task(_periodic_flush())
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
