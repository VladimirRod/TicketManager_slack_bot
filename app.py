from slack_bolt.async_app import AsyncApp
from slack_bolt.adapter.socket_mode.async_handler import AsyncSocketModeHandler
from dateutil import tz
import dateutil.parser
import random
import string
import datetime
import json
import gspread

app = AsyncApp(token="xoxb-*************-*************-************************")  # SLACK_BOT_TOKEN

service_account = gspread.service_account(filename='credentials.json')  # Google Sheet API token
sheet_id = service_account.open_by_key('********************************************')  # Sheet_id
worksheet = sheet_id.sheet1

sheet_mentors = service_account.open_by_key('********************************************')  # Sheet_id
channel_mentor = sheet_mentors.sheet1


@app.event("message")
async def create_ticket(client, message, ack, body):
    if "parent_user_id" not in json.dumps(body["event"]):
        user = message.get('user')
        text = message.get('text')
        channel_id = message.get('channel')
        thr_ts = message.get('ts')
        cell = channel_mentor.find(channel_id)
        mentor = channel_mentor.cell(cell.row, cell.col + 1).value
        # Преобразуем thr_ts из timestamp в UTC и приводим к зоне МСК
        message_ts = dateutil.parser.isoparse(
            datetime.datetime.fromtimestamp(round(float(thr_ts))).isoformat()).astimezone(tz.gettz('Europe/Moscow'))
        messagets2json = json.dumps(message_ts, default=str).replace('"', '')
        ticket_id = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))

        await ack()
        await client.chat_postMessage(
            thread_ts=thr_ts,
            channel=channel_id,
            text=f"{text}",
            blocks=[
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": f"Привет! Твоё обращение зарегистрировано и ему присвоен номер {ticket_id}. Призываю <@{mentor}>"
                    }
                },
                {
                    "type": "actions",
                    "elements": [
                        {
                            "type": "button",
                            "text": {
                                "type": "plain_text",
                                "text": "Взять в работу",
                            },
                            "style": "primary",
                            "action_id": "in_work",
                            "value": f"{ticket_id}{channel_id}",
                        }
                    ]

                }
            ]
        )
        offset = datetime.timezone(datetime.timedelta(hours=3))
        created_at = datetime.datetime.now(offset).replace(microsecond=0)
        created2json_ts = json.dumps(created_at, default=str).replace('"', '')
        # Записываем в таблицу номер тикета, время сообщения пользователя, айди канала, текст и время создания тикета
        worksheet.append_row([ticket_id, messagets2json, channel_id, text, created2json_ts])


@app.action("in_work")
async def in_work_progress(action, client, ack, body):
    channel_id = action["value"][6:]
    ticket_id = action["value"][:6]
    text_message = body["message"]["text"]
    timestamp = body["container"]["message_ts"]
    await ack()
    await client.chat_update(
        channel=body["channel"]["id"],
        ts=timestamp,
        text=channel_id,
        blocks=[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": "Исполнитель начал работу над задачей"
                }
            },
            {
                "type": "actions",
                "elements": [
                    {
                        "type": "button",
                        "text": {
                            "type": "plain_text",
                            "text": "Закрыть обращение",
                        },
                        "style": "danger",
                        "action_id": "close_ticket",
                        "value": f"{ticket_id}{timestamp}"
                    }
                ]
            }
        ]
    )
    offset = datetime.timezone(datetime.timedelta(hours=3))
    in_work_ts = datetime.datetime.now(offset).replace(microsecond=0)
    inworkts2json = json.dumps(in_work_ts, default=str).replace('"', '')

    # Записываем в таблицу время взятия в работу
    cell = worksheet.find(ticket_id)
    await ack()
    if cell:
        worksheet.update_cell(cell.row, 6, inworkts2json)


@app.action("close_ticket")
async def close_ticket(action, client, ack, body):
    ticket_id = action["value"][:6]
    msms_ts = action["value"][6:]
    thread_ts = body["container"]["thread_ts"]
    channel_id = body["container"]["channel_id"]
    text_message = str(body["message"]["text"])
    await ack()
    await client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "view_1",
            "private_metadata": f"{thread_ts}{channel_id}{msms_ts}{ticket_id}",
            "title": {
                "type": "plain_text",
                "text": "Закрытие обращения",
            },
            "submit": {
                "type": "plain_text",
                "text": "Закрыть",
            },
            "close": {
                "type": "plain_text",
                "text": "Отмена",
            },
            "blocks": [
                {
                    "type": "input",
                    "block_id": "type",
                    "element": {
                        "type": "static_select",
                        "placeholder": {
                            "type": "plain_text",
                            "text": "Выберите тип обращения",
                        },
                        "action_id": "type",
                        "options": [
                            {
                                "text": {"type": "plain_text",
                                         "text": "Вопрос по проверенному домашнему заданию/проекту"},
                                "value": "option_1",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Как делать домашку"},
                                "value": "option_2",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Непонятная формулировка в модуле"},
                                "value": "option_3",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Просто что-то не работает"},
                                "value": "option_4",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Ошибка в модуле"},
                                "value": "option_5",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Устаревшая информация в курсе"},
                                "value": "option_6",
                            },
                            {
                                "text": {"type": "plain_text",
                                         "text": "Не работает стороннее сервис/приложение, студент не знает, что делать"},
                                "value": "option_7",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Где взять дополнительную информацию"},
                                "value": "option_8",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Ошибочный запрос"},
                                "value": "option_9",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Консультация"},
                                "value": "option_10",
                            },
                        ],
                    },
                    "label": {
                        "type": "plain_text",
                        "text": "Тип обращения",
                    },
                },
                {
                    "type": "input",
                    "block_id": "tag",
                    "element": {
                        "type": "static_select",
                        "placeholder": {
                            "type": "plain_text",
                            "text": "Укажите тег для обращения",
                        },
                        "action_id": "tag",
                        "options": [
                            {
                                "text": {"type": "plain_text", "text": "Решено"},
                                "value": "tag_1",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Не делаем"},
                                "value": "tag_2",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Передано менторам/админам"},
                                "value": "tag_3",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Передано на исправление/внесение изменений"},
                                "value": "tag_4",
                            },
                        ],
                    },
                    "label": {
                        "type": "plain_text",
                        "text": "Тег",
                    },
                },
                {
                    "type": "input",
                    "block_id": "notes",
                    "optional": True,
                    "element": {
                        "type": "plain_text_input",
                        "multiline": True,
                        "action_id": "notes",
                        "placeholder": {
                            "type": "plain_text",
                            "text": "Краткий комментарий по сути проблемы",
                        },
                    },
                    "label": {
                        "type": "plain_text",
                        "text": "Заметки",
                    },
                },
            ],
        })
    closedby = body["user"]["username"]

    # Записываем в таблицу логин, закрывшего обращение
    cell = worksheet.find(ticket_id)
    await ack()
    if cell:
        worksheet.update_cell(cell.row, 8, closedby)


@app.view("view_1")
async def handle_view(ack, body, view, client):
    channel_id = body["view"]["private_metadata"][17:28]
    msms_ts = body["view"]["private_metadata"][28:45]
    thread_ts = body["view"]["private_metadata"][:17]
    ticket_id = body["view"]["private_metadata"][45:]
    type = view["state"]["values"]["type"]["type"]["selected_option"]["text"]["text"]
    tag = view["state"]["values"]["tag"]["tag"]["selected_option"]["text"]["text"]
    notes = view["state"]["values"]["notes"]["notes"]["value"]
    await ack()
    await client.chat_postMessage(
        thread_ts=thread_ts,
        channel=channel_id,
        blocks=[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": f"Обращение с номером {ticket_id} закрыто"
                }
            },
            {
                "type": "actions",
                "elements": [
                    {
                        "type": "button",
                        "text": {
                            "type": "plain_text",
                            "text": "Переоткрыть обращение",
                        },
                        "style": "primary",
                        "action_id": "reopen_ticket",
                        "value": f"{ticket_id}"
                    }
                ]

            }
        ]
    )
    # Собираем время закрытия тикета, приводим к зоне МСК и убираем микросекунды
    offset = datetime.timezone(datetime.timedelta(hours=3))
    closed_at = datetime.datetime.now(offset).replace(microsecond=0)
    closed2json_ts = json.dumps(closed_at, default=str).replace('"', '')

    await ack()
    await client.chat_update(
        channel=channel_id,
        ts=msms_ts,
        blocks=[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": "Исполнитель начал работу над задачей"
                }
            }
        ]
    )

    # Записываем в таблицу
    cell = worksheet.find(ticket_id)
    if cell:
        worksheet.update_cell(cell.row, 7, closed2json_ts)
        worksheet.update_cell(cell.row, 9, tag)
        worksheet.update_cell(cell.row, 10, type)
        worksheet.update_cell(cell.row, 11, notes)


@app.action("reopen_ticket")
async def reopen(ack, body, client, action):
    ticket_id = action["value"]
    channel_id = body["container"]["channel_id"]
    reopen_thread_ts = body["container"]["message_ts"]
    await ack()
    await client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "view_2",
            "private_metadata": f"{ticket_id}{reopen_thread_ts}{channel_id}",
            "title": {
                "type": "plain_text",
                "text": "Переоткрытие обращения",
            },
            "submit": {
                "type": "plain_text",
                "text": "Переоткрыть",
            },
            "close": {
                "type": "plain_text",
                "text": "Отмена",
            },
            "blocks": [
                {
                    "type": "input",
                    "block_id": "type",
                    "element": {
                        "type": "static_select",
                        "placeholder": {
                            "type": "plain_text",
                            "text": "Выберите тип обращения",
                        },
                        "action_id": "type",
                        "options": [
                            {
                                "text": {"type": "plain_text", "text": "Не хватило информации в ответе ментора"},
                                "value": "option_1",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Хочу уточнить еще кое-что по теме"},
                                "value": "option_2",
                            },
                        ],
                    },
                    "label": {
                        "type": "plain_text",
                        "text": "Тип обращения",
                    },
                },
                {
                    "type": "input",
                    "block_id": "notes",
                    "optional": True,
                    "element": {
                        "type": "plain_text_input",
                        "multiline": True,
                        "action_id": "notes",
                        "placeholder": {
                            "type": "plain_text",
                            "text": "Краткий комментарий по сути проблемы",
                        },
                    },
                    "label": {
                        "type": "plain_text",
                        "text": "Заметки",
                    },
                },
            ],
        })


@app.view("view_2")
async def reopen_view(ack, body, view, client):
    reopen_type = view["state"]["values"]["type"]["type"]["selected_option"]["text"]["text"]
    reopen_notes = view["state"]["values"]["notes"]["notes"]["value"]
    ticket_id = body["view"]["private_metadata"][:6]
    reopen_thread_ts = body["view"]["private_metadata"][6:23]
    channel_id = body["view"]["private_metadata"][23:]
    offset = datetime.timezone(datetime.timedelta(hours=3))
    reopents = datetime.datetime.now(offset).replace(microsecond=0)
    reopen2json_ts = str(json.dumps(reopents, default=str).replace('"', ''))
    cell = worksheet.find(ticket_id)
    await ack()
    if cell:
        worksheet.update_cell(cell.row, 12, reopen2json_ts)
        worksheet.update_cell(cell.row, 13, str(reopen_type))
        worksheet.update_cell(cell.row, 14, str(reopen_notes))
    await ack()
    await client.chat_update(
        channel=channel_id,
        ts=reopen_thread_ts,
        blocks=[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": "Обращение переоткрыто"
                }
            },
            {
                "type": "actions",
                "elements": [
                    {
                        "type": "button",
                        "text": {
                            "type": "plain_text",
                            "text": "Закрыть обращение",
                        },
                        "style": "danger",
                        "action_id": "close_reopen_ticket",
                        "value": f"{ticket_id}"
                    }
                ]

            }
        ]
    )


@app.action('close_reopen_ticket')
async def close_reopen_ticket(ack, client, body, action):
    ticket_id = action["value"]
    reopen_ms_ts = body["container"]["message_ts"]
    channel_id = body["container"]["channel_id"]
    await ack()
    await client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "view_3",
            "private_metadata": f"{ticket_id}{reopen_ms_ts}{channel_id}",
            "title": {
                "type": "plain_text",
                "text": "Закрытие обращения",
            },
            "submit": {
                "type": "plain_text",
                "text": "Закрыть",
            },
            "close": {
                "type": "plain_text",
                "text": "Отмена",
            },
            "blocks": [
                {
                    "type": "input",
                    "block_id": "type",
                    "element": {
                        "type": "static_select",
                        "placeholder": {
                            "type": "plain_text",
                            "text": "Выберите тип обращения",
                        },
                        "action_id": "type",
                        "options": [
                            {
                                "text": {"type": "plain_text",
                                         "text": "Вопрос по проверенному домашнему заданию/проекту"},
                                "value": "option_1",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Как делать домашку"},
                                "value": "option_2",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Непонятная формулировка в модуле"},
                                "value": "option_3",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Просто что-то не работает"},
                                "value": "option_4",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Ошибка в модуле"},
                                "value": "option_5",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Устаревшая информация в курсе"},
                                "value": "option_6",
                            },
                            {
                                "text": {"type": "plain_text",
                                         "text": "Не работает стороннее сервис/приложение, студент не знает, что делать"},
                                "value": "option_7",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Где взять дополнительную информацию"},
                                "value": "option_8",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Ошибочный запрос"},
                                "value": "option_9",
                            },
                            {
                                "text": {"type": "plain_text", "text": "Консультация"},
                                "value": "option_10",
                            },
                        ],
                    },
                    "label": {
                        "type": "plain_text",
                        "text": "Тип обращения",
                    },
                },
                {
                    "type": "input",
                    "block_id": "tag",
                    "element": {
                        "type": "static_select",
                        "placeholder": {
                            "type": "plain_text",
                            "text": "Укажите тег для обращения",
                        },
                        "action_id": "tag",
                        "options": [
                            {
                                "text": {"type": "plain_text", "text": "Вопрос закрыт"},
                                "value": "tag_1",
                            },
                        ],
                    },
                    "label": {
                        "type": "plain_text",
                        "text": "Тег",
                    },
                },
                {
                    "type": "input",
                    "block_id": "notes",
                    "optional": True,
                    "element": {
                        "type": "plain_text_input",
                        "multiline": True,
                        "action_id": "notes",
                        "placeholder": {
                            "type": "plain_text",
                            "text": "Краткий комментарий по сути проблемы",
                        },
                    },
                    "label": {
                        "type": "plain_text",
                        "text": "Заметки",
                    },
                },
            ],
        })


@app.view("view_3")
async def close_2(ack, client, body, view):
    ticket_id_id = body["view"]["private_metadata"][:6]
    reopen_ms_ts = body["view"]["private_metadata"][6:23]
    channel_id = body["view"]["private_metadata"][23:]
    close2_type = view["state"]["values"]["type"]["type"]["selected_option"]["text"]["text"]
    close2_tag = view["state"]["values"]["tag"]["tag"]["selected_option"]["text"]["text"]
    close2_notes = view["state"]["values"]["notes"]["notes"]["value"]
    await ack()
    await client.chat_update(
        channel=channel_id,
        ts=reopen_ms_ts,
        blocks=[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": f"Обращение с номером {ticket_id_id} закрыто"
                }
            }
        ]
    )
    offset = datetime.timezone(datetime.timedelta(hours=3))
    close2ts = datetime.datetime.now(offset).replace(microsecond=0)
    close22json_ts = str(json.dumps(close2ts, default=str).replace('"', ''))

    cell = worksheet.find(ticket_id_id)
    await ack()
    if cell:
        worksheet.update_cell(cell.row, 15, close22json_ts)
        worksheet.update_cell(cell.row, 16, close2_type)
        worksheet.update_cell(cell.row, 17, close2_tag)
        worksheet.update_cell(cell.row, 18, close2_notes)


async def main():
    handler = AsyncSocketModeHandler(app,"xapp-*-***********-*************-****************************************************************") # APP_LEVEL_TOKEN
    await handler.start_async()


if __name__ == "__main__":
    import asyncio

    asyncio.run(main())
