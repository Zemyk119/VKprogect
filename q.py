import asyncio
import os

import aiosqlite
from openpyxl import load_workbook


async def validator(address):
    book = load_workbook(address)
    sheet = book['Лист1']
    upload_list = []
    for i in range(1, 146, 3):
        # day, number, mod = None, None, None
        if sheet['A' + str(i)].value:
            day = sheet['A' + str(i)].value
        if sheet['B' + str(i)].value:
            number = sheet['B' + str(i)].value
        if sheet['C' + str(i)].value:
            mod = sheet['C' + str(i)].value
        if day and number and mod and sheet['D' + str(i)].value and sheet['D' + str(i + 1)].value and sheet[
            'D' + str(i + 2)].value:
            upload_list.append([day, number, mod, sheet['D' + str(i)].value, sheet['D' + str(i + 1)].value, sheet[
                'D' + str(i + 2)].value, 0])

        if sheet['F' + str(i)].value:
            day1 = sheet['F' + str(i)].value
        if sheet['G' + str(i)].value:
            number1 = sheet['G' + str(i)].value
        if sheet['H' + str(i)].value:
            mod1 = sheet['H' + str(i)].value
        if day1 and number1 and mod1 and sheet['I' + str(i)].value and sheet['I' + str(i + 1)].value and sheet[
            'I' + str(i + 2)].value:
            upload_list.append([day1, number1, mod1, sheet['I' + str(i)].value, sheet['I' + str(i + 1)].value, sheet[
                'I' + str(i + 2)].value, 1])

    connection = await aiosqlite.connect('all_users.db')
    command = await connection.cursor()
    for inf in upload_list:
        await command.execute(
            '''
            UPDATE schedule SET day = ?, number = ?, type = ? , name = ?, teacher = ?, place = ?, parity = ?, 
            subgroup = ?, group_number = ?, course = ?;
            ''',
            (inf[0], inf[1], inf[2], inf[3], inf[4],
             inf[5], inf[6], 9, 9, 9)
        )
    await connection.commit()
    await command.close()
    await connection.close()
    os.remove(address)


asyncio.run(validator('excel_file/310877068.xlsx'))
