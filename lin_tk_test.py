'''
1. 도서목록 데이터가 저장된 엑셀파일을 불러온다
2. row, column으로 구분하여 도서 목록을 분류한다.
3. 분류된 도서에 대출가능상태를 설정, 변경 할 수 있게 한다 -> 도서 검색 엔진, 대출 엔진, 반납 엔진
4. 전체 도서목록을 대출가능상태와 같이 GUI로 구현한다.
5. print() 함수는 백엔드 작업시에만 필요, GUI진행이 완료되면 삭제추천.
'''

import openpyxl
from tkinter import *
def refresh():
    '''1. 도서목록 데이터가 저장된 엑셀파일을 불러온다'''
    global book_excel_file
    book_excel_file = openpyxl.load_workbook('book_list.xlsx')
    global file_name
    file_name = 'book_list.xlsx'
    global list_sheet
    list_sheet = book_excel_file.worksheets[0]

    '''2. row, column으로 구분하여 도서 목록을 리스트에 저장한다.'''
    global  book_list
    book_list = []
    #분류번호, 제목, 대출가능 순으로 dict type으로 생성후 list에 저장
    for row in list_sheet.rows:
        data = {}
        data['number'] = row[0].value
        data['name'] = row[1].value
        data['loan'] = row[2].value
        book_list.append(data)

    #엑셀에 row1은 불필요하므로 삭제
    del book_list[0]
    return
################################
#데이터가 정상적으로 만들어졌는지 확인
#for data in book_list:
#    print(data)
#    print(data['name'])
################################

'''도서 검색 엔진'''

#제목으로 검색
def search_engine(book_name):
    check = None
    for data in book_list:
        if book_name in data['name']:
            output.insert(CURRENT, f"********************\n")
            output.insert(CURRENT, f"{data['number']}\n")
            output.insert(CURRENT, f"{data['name']}\n")
            #print(data['number'])
            #print(data['name'])
            if data['loan'] == 1:
                output.insert(CURRENT, '대출가능\n')
                output.insert(CURRENT, f"********************\n")
                print('대출가능')
            else:
                output.insert(CURRENT, '이미 대출되어있는 도서입니다.\n')
                output.insert(CURRENT, "********************\n")
                print('이미 대출되어있는 도서입니다.')
            print("")
            check = True #중복되는 제목의 책을 찾고 책을 아예 찾지 못했을 경우 안내문을 print해주기 위한 변수
    if check == True:
        return
    else:
        output.insert(CURRENT, '찾으려는 책이 없습니다. 제목을 확인해주세요\n\n')
        print("찾으려는 책이 없습니다. 제목을 확인해주세요")


#분류번호로 검색
def search_engine_number(number):
    for data in book_list:
        if number == data['number']:
            output.insert(CURRENT, f"********************\n")
            output.insert(CURRENT, f"{data['number']}\n")
            output.insert(CURRENT, f"{data['name']}\n")
            print(data['number'])
            print(data['name'])
            if data['loan'] == 1:
                print('대출가능')
                output.insert(CURRENT, '대출가능\n')
                output.insert(CURRENT, f"********************\n")
            else:
                output.insert(CURRENT, '이미 대출되어있는 도서입니다.\n')
                output.insert(CURRENT, f"********************\n")
                print('이미 대출되어있는 도서입니다.')
            return
    output.insert(CURRENT, '찾으려는 책이 없습니다. 분류번호를 확인해주세요\n\n')
    print('찾으려는 책이 없습니다. 분류번호를 확인해주세요')

#실제 검색 구현
def search():
    refresh()
    output.delete('1.0', END)
    user_entry = entry.get()
    #user_type = input('책 제목 또는 분류번호를 입력하세요 :')
    if user_entry[0] == '2':
        search_engine_number(user_entry)
    else:
        search_engine(user_entry)

'''도서 대출 엔진'''
def book_loan_engine(book_name):
    for data in book_list:
        if book_name in data['name']:
            if data['loan'] == 1:
#                print(f"{data['name']} 대출 완료")
                output.insert(CURRENT, f"********************\n")
                output.insert(CURRENT, f"{data['number']}\n")
                output.insert(CURRENT, f"{data['name']}\n")
                print(data['number'])
                print(data['name'])
                output.insert(CURRENT, '대출완료\n')
                output.insert(CURRENT, f"********************\n")
                print('대출 완료')
                row_number = book_list.index(data)
                input_cell = row_number + 2
                list_sheet[f'C{input_cell}'] = 0
                book_excel_file.save(filename = file_name) #변경된 데이터 엑셀에 다시 저장! 중요!!
                return
            else:
                output.insert(CURRENT, '이미 대출되어있는 도서입니다.\n\n')
                print('이미 대출되어있는 도서입니다.')
                return
    output.insert(CURRENT, '대출하려는 책이 없습니다. 제목을 확인해주세요.네임\n\n')
    print('대출하려는 책이 없습니다. 제목을 확인해주세요')

def book_loan_engine_number(number):
    for data in book_list:
        if number == data['number']:
            if data['loan'] == 1:
                output.insert(CURRENT, f"********************\n")
                output.insert(CURRENT, f"{data['number']}\n")
                output.insert(CURRENT, f"{data['name']}\n")
                print(data['number'])
                print(data['name'])
                output.insert(CURRENT, '대출완료\n')
                output.insert(CURRENT, f"********************\n")
                print('대출 완료\n')
                row_number = book_list.index(data)
                input_cell = row_number + 2
                list_sheet[f'C{input_cell}'] = 0
                book_excel_file.save(filename = file_name)
                return
            else:
                output.insert(CURRENT, '이미 대출되어있는 도서입니다.\n')
                output.insert(CURRENT, f"********************\n")
                print('이미 대출되어있는 도서입니다.')
                return
    output.insert(CURRENT, '대출하려는 책이 없습니다. 제목을 확인해주세요.넘버\n\n')
    print('대출하려는 책이 없습니다. 제목을 확인해주세요')


#실제 대출 구현
def loan():
    refresh()
    output.delete('1.0', END)
    #user_type = input('책 제목을 입력하세요 :')
    user_entry = entry.get()
    if user_entry[0] == '2':
        book_loan_engine_number(user_entry)
    else:
        book_loan_engine(user_entry)

'''도서 반납 엔진'''
def ban_nap_engine(number):
    for data in book_list:
        if number == data['number']:
            if data['loan'] == 0:
                output.insert(CURRENT, f"********************\n")
                output.insert(CURRENT, f"{data['number']}\n")
                output.insert(CURRENT, f"{data['name']}\n")
                print(data['number'])
                print(data['name'])
                output.insert(CURRENT, '반납완료\n\n')
                output.insert(CURRENT, f"********************\n")
                print('반납 완료')
                row_number = book_list.index(data)
                input_cell = row_number + 2
                list_sheet[f'C{input_cell}'] = 1
                book_excel_file.save(filename=file_name)
                return
            else:
                output.insert(CURRENT, '이미 반납되어있는 도서입니다.\n\n')
                print('이미 반납되어있는 도서입니다.')
                return
    output.insert(CURRENT, '반납하려는 책은 도서관에 없는 책입니다. 목록을 확인해주세요.\n\n')
    print('반납하려는 책은 도서관에 없는 책입니다. 목록을 확인해주세요.')

#실제 도서 반납 구현
def ban_nap():
    refresh()
    output.delete('1.0', END)
    #user_type = input('분류 번호를 입력해주세요 (예)20-도B-01 : ')
    user_entry = entry.get()
    ban_nap_engine(user_entry)

'''GUI구현'''
window = Tk()
window.title("Franken Stein")
window.geometry('640x400+10+10')

entry = Entry(window, width=70)
#entry.place(x=0, y=0)
entry.grid(row=0, column=0)

search_button = Button(window, text='검색', command=search)
search_button.grid(row=0, column=1)

loan_button = Button(window, text='대출', command=loan)
loan_button.grid(row=1, column=1)

ban_nap_button = Button(window, text='반납', command=ban_nap)
ban_nap_button.grid(row=2, column=1)

output = Text(window, width=70)
output.grid(row=1, column=0)
#output.place(x=0, y=50)

window.mainloop()

