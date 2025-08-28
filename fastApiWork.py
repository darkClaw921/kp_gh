import time
# import redisWork
from fastapi import FastAPI, Request, Form
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi import BackgroundTasks
from pydantic import BaseModel
from pprint import pprint  
from datetime import datetime

import asyncio
from dotenv import load_dotenv
import os
# from loguru import logger

# log = logger()
# from loguru import logger
# logger.add("logs/fastApi_{time}.log",format="{time} - {level} - {message}", rotation="100 MB", retention="10 days", level="DEBUG")

load_dotenv()

PORT = os.getenv('PORT')
app = FastAPI(title='GAB API', description='Генерация КП для сделки')

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Adjust this as needed
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)



@app.post('/event')
async def update_event(request: Request):
    """Обновление сущности"""
    form_data = await request.form()
    data = {key: form_data[key] for key in form_data.keys()}
    pprint(data)

    event = data.get('event')
    print(f"{event=}")
    
    # if event == 'ONCALENDARENTRYADD':
    #     eventID = data['data[id]']
    #     print(f"{eventID=}")
    #     event = await get_calendar_event(eventID)
    #     # pprint(event)
    #     await create_billing_for_event(event=event)

    # elif event == 'ONCALENDARENTRYUPDATE':
    #     eventID = data['data[id]']
    #     print(f"{eventID=}")
    #     event = await get_calendar_event(eventID)
    #     # pprint(event)
    #     await update_billing_for_event(event=event)
        

    # elif event == 'ONTASKUPDATE':
    #     taskID = data['data[FIELDS_BEFORE][ID]']
    #     await create_billing_for_task(taskID=taskID)

    return JSONResponse(content={'message': 'OK'})

# async def update_event_local(request: dict):
#     """Обновление сущности"""
#     # form_data = await request.form()
#     # data = {key: form_data[key] for key in form_data.keys()}
#     # FIELDS_POLE_OLD="UF_CRM_1752673288234"

#     data= request
#     pprint(data)

#     event = data.get('event')
#     print(f"{event=}")
    
#     if event == 'ONCALENDARENTRYADD':
#         eventID = data['data[id]']
#         print(f"{eventID=}")
#         event = await get_calendar_event(eventID)
#         # pprint(event)
#         billingID= await create_billing_for_event(event=event)
#         await add_billing_to_event(eventID=eventID, billingID=billingID)

#     elif event == 'ONCALENDARENTRYUPDATE':
#         eventID = data['data[id]']
#         print(f"{eventID=}")
#         event = await get_calendar_event(eventID)
#         # pprint(event)
#         await update_billing_for_event(event=event)
        

#     elif event == 'ONTASKUPDATE':
#         taskID = data['data[FIELDS_BEFORE][ID]']
#         await create_billing_for_task(taskID=taskID)

#     elif event == 'ONCRMCONTACTUPDATE' or event =='ONCRMCONTACTADD':
        
#         print("данные из запроса ", data)
#         contactID = data['data[FIELDS][ID]']
#         contact = await get_contact(contactID=contactID)
#         # if event == 'ONCRMCONTACTADD':
#         print(f"найден контакт {contact}")
#         print('запущено создание нового контакта')
#         if isinstance(contact, list):
#             phone=contact[0]['PHONE'][-1]['VALUE']
#         else:
#             phone=contact['PHONE'][-1]['VALUE']
#         print('поиск дубликатов для номера ', phone)

#         duplicateContacts=await find_duplicate_contacts(phone=phone)
#         # pprint(duplicateContacts)
#         print('найдены дубликаты ', duplicateContacts)
#         if duplicateContacts not in [None, []]:
#             newContactID=max(duplicateContacts) 
#             for duplicateContactID in duplicateContacts:
#                 if duplicateContactID != newContactID:
#                     await merge_contacts(oldContactID=duplicateContactID, newContactID=newContactID)
#             contact=await get_contact(contactID=newContactID)
#                     # return JSONResponse(content={'message': 'OK'})
            

#         print(f"работает с контактом {contact=}")
#         phone = contact.get('PHONE')
#         print(f"найден телефон {phone=}")
#         if phone:
#             phone = phone[-1]['VALUE']
#             phoneOld=contact.get(FIELDS_POLE_OLD)
            
#             if phone != phoneOld:
#                 deals = await get_deal_for_contact(contactID)
#                 print(f"найдены сделки по контакту {contactID=} {deals=}")
                
#                 deal=deals
#                 # deal['STAGE_ID']='NEW'
                

#                 phones=await delete_old_phones(phones=contact['PHONE'], newPhone=phone)
                
#                 fields={
#                     FIELDS_POLE_OLD: phone,
#                     'PHONE':phones
#                 }
#                 print(f"обновляем контакт {fields=}")

#                 await update_contact(contactID=contactID, fields=fields)
#                 print("обновил контакт", contactID)
                
#                 print(f"обновляем сделку  переносим в другую стадию {deal=}")
#                 print(f"статус сделки {deal['STAGE_SEMANTIC_ID']=}")
#                 print(f"категория сделки {deal['CATEGORY_ID']=}")
#                 print(f"статус сделки {deal['STAGE_ID']=}")
#                 print(f"проверка условия {(deal['STAGE_SEMANTIC_ID'] == 'P')=}")
#                 print(f"проверка условия {(deal['CATEGORY_ID'] in ['1', '0'])=}")
#                 print(f"проверка условия {(deal['STAGE_SEMANTIC_ID'] == 'P' and (deal['CATEGORY_ID'] in ['1', '0']))=}")

#                 if deal['STAGE_SEMANTIC_ID'] == 'P' and (deal['CATEGORY_ID'] in ['1', '0']):
#                     fields={
#                         # 'STAGE_ID': 'NEW'
#                         'STAGE_ID': 'lv6nyzyhbmpl-197-1cl1b',
#                         # 'CATEGORY_ID': '0'
#                     }
#                     print("данные для обновления сделки", fields)

#                     if deal['CATEGORY_ID'] == '1':
#                         print("перенос сделки в стадию 0")
#                         await update_deal_category(dealID=deal['ID'], categoryId='0')
                    
#                     await update_deal(dealID=deal['ID'], fields=fields)
#                     print("сделка перенесена в другую стадию", deal['ID'])
            



#     return JSONResponse(content={'message': 'OK'})



if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host='0.0.0.0', port=int(PORT), log_level="info")