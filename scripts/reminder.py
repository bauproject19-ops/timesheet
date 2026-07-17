"""Ежедневное личное напоминание сотрудникам (17:30 МСК) + ругань за вчера,
если вчера был рабочий день и человек не заполнил часы. В дни отпуска
сотруднику ничего не шлём."""
from datetime import date, datetime, timedelta, timezone

from bitrix_common import (
    REMINDER_USER_IDS, FORM_URL, is_day_off, get_all_items,
    record_date, vacationing_users_on, send_personal,
)


def main():
    now_msk = datetime.now(timezone.utc) + timedelta(hours=3)
    today = now_msk.date()
    yesterday = today - timedelta(days=1)

    items = get_all_items()
    vac_today = vacationing_users_on(today, items)
    vac_yesterday = vacationing_users_on(yesterday, items)

    def has_record_on(user_id: int, d: date) -> bool:
        return any(
            it["assignedById"] == user_id and record_date(it.get("title", "")) == d
            for it in items
        )

    day_off_today = is_day_off(today)
    day_off_yesterday = is_day_off(yesterday)

    results = []
    for uid in REMINDER_USER_IDS:
        if uid in vac_today:
            results.append((uid, "skipped: отпуск сегодня"))
            continue

        scold = False
        if not day_off_yesterday and uid not in vac_yesterday:
            scold = not has_record_on(uid, yesterday)

        # Персональная ссылка с ?uid= — форма на неё закрепляет сотрудника и
        # блокирует выпадающий список, чтобы часы случайно не ушли на другого
        # человека (было: общая ссылка на всех + запоминание выбора в
        # localStorage браузера — реальный кейс, когда часы улетели не туда).
        form_link = f"{FORM_URL}?uid={uid}"

        msg = ""
        if scold:
            msg += (
                f"Обрати внимание: за {yesterday.strftime('%d.%m.%Y')} (рабочий день) "
                f"часы так и не заполнены. Пожалуйста, занеси их задним числом "
                f"(выбери нужную дату в поле \"Дата\" формы): {form_link}\n\n"
            )
        if day_off_today:
            msg += (
                f"Сегодня {today.strftime('%d.%m.%Y')} — выходной/праздник. Норма 0ч, "
                f"заполнять не обязательно. Если всё же работал(а) сегодня — занеси часы: {form_link}"
            )
        else:
            msg += (
                f"Не забудь заполнить часы за {today.strftime('%d.%m.%Y')} "
                f"(рабочий день, норма 8ч). Форма: {form_link}"
            )

        send_personal(uid, msg)
        results.append((uid, "ok"))

    for uid, status in results:
        print(f"{uid}: {status}")


if __name__ == "__main__":
    main()
