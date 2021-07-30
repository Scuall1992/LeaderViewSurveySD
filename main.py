state_indexes = [1,2,3,4,11,14,20]
intent_indexes = [6,7,8,9,10,12,16]
relationship_indexes = [5,13,14,17,18,19,22,23,24,25]
awareness_indexes = [10,11,15,17,18,19,21]


def send_email(who, data):
    import smtplib, ssl
    from email.mime.text import MIMEText
    from email.header    import Header

    smtp_server = "smtp.gmail.com"
    sender_email = "scuall1992@gmail.com"  # Enter your address
    receiver_email = who  # Enter receiver address
    password = 'dhjsrvwexoxeajms'
    message = data
    

    smtp_host = 'smtp.gmail.com'       # google
    
    msg = MIMEText(data, 'plain', 'utf-8')
    msg['Subject'] = Header('Результаты теста по спиральной динамике', 'utf-8')
    msg['From'] = sender_email
    msg['To'] = who

    s = smtplib.SMTP(smtp_host, 587, timeout=10)
    s.set_debuglevel(1)
    try:
        s.starttls()
        s.login(sender_email, password)
        s.sendmail(msg['From'], who, msg.as_string())
    finally:
        s.quit()


def get_sira(people): # 10 - max

    people = people["sira"]

    S = sum([people[i-1] for i in state_indexes]) / 6.3
    I = sum([people[i-1] for i in intent_indexes]) / 6.3
    R = sum([people[i-1] for i in relationship_indexes]) / 9
    A = sum([people[i-1] for i in awareness_indexes]) / 6.3

    F = (S+I+R+A)/4


    return list(map(lambda x: round(x, 2), (S, I, R, A, F)))



purple = [4,9,19,24]
red = [2,8,15,23]
blue = [1,12,17,27]
orange = [6,13,20,22]
green = [5,11,18,26]
yellow = [7,10,16,25]
turquoise = [3,14,21,28]


def get_mindset(people): # 60 - max

    people = people["mindset"]

    b = sum([people[i-1] for i in purple if people[i-1] != ""])
    c = sum([people[i-1] for i in red if people[i-1] != ""])
    d = sum([people[i-1] for i in blue if people[i-1] != ""])
    e = sum([people[i-1] for i in orange if people[i-1] != ""])
    f = sum([people[i-1] for i in green if people[i-1] != ""])
    g = sum([people[i-1] for i in yellow if people[i-1] != ""])
    h = sum([people[i-1] for i in turquoise if people[i-1] != ""])

    return [int(i) if i != "" else 0 for i in [b,c,d,e,f,g,h]]
    
def get_duress_mindset(people): # 15 - max

    people = people["duress"]

    b = people[2] #purple
    c = people[0] #red
    d = people[3] #blue
    e = people[1] #orange
    f = people[4] #green
    g = people[5] #yellow
    h = people[6] #turquoise

    return [int(i) if i != "" else 0 for i in [b,c,d,e,f,g,h]]


def get_change_state(people): # 12 - max
    
    people = people["change_state"]

    alpha = [0,9]
    beta = [1,6]
    gamma = [2,5]
    delta = [3,7]
    new_alpha = [4,8]

    a = sum([people[i-1] for i in alpha])
    b = sum([people[i-1] for i in beta])
    c = sum([people[i-1] for i in gamma])
    d = sum([people[i-1] for i in delta])
    e = sum([people[i-1] for i in new_alpha])

    
    return list(map(int, [a,b,c,d,e]))


import xlrd 

book = xlrd.open_workbook("3.xlsx")

sheet = book.sheet_by_index(0)

peoples = []

# sirafa
for rownum in range(sheet.nrows):
    row = sheet.row_values(rownum)
    

    start = 4
    end = 29
    
    info = row[:start]

    sira = row[start:end]

    start = end
    end += 7*4
    mindset = row[start:end]

    start = end 
    end += 7
    duress = row[start:end]
    
    start = end 
    end += 10
    change_state = row[start:end]
    
    start = end 
    end += 6
    answers = row[start:end]

    peoples.append({"info":info, "sira":sira, "mindset":mindset, 
                    "duress":duress, "change_state":change_state, "answers":answers})



header = peoples[0]

import os 

FOLDER = "results"

if not os.path.exists(FOLDER):
    os.mkdir(FOLDER)

for people in peoples[1:]:
    filename = os.path.join(FOLDER, f"{people['info'][1]} {people['info'][2]}.txt")

    with open(filename, "w", encoding="utf-8") as f:
        sira = get_sira(people)

        #f.write("SIRAFA:\n\n")
        # print(filename)
        # print(f'Состояние: {sira[0]}/10\n')
        # print(f'Намерение: {sira[1]}/10\n')
        # print(f'Взаимоотношения: {sira[2]}/10\n')
        # print(f'Осознанность: {sira[3]}/10\n')
        # print(f'Гибкость: {sira[4]}/10\n\n')
        # print()

        mindset = get_mindset(people)

        f.write("Мировоззрение:\n\n")
        f.write(f'Фиолетовый: {mindset[0]}/60 {round(mindset[0]*100/60,2)}%\n')
        f.write(f'Красный: {mindset[1]}/60 {round(mindset[1]*100/60,2)}%\n')
        f.write(f'Синий: {mindset[2]}/60 {round(mindset[2]*100/60,2)}%\n')
        f.write(f'Оранжевый: {mindset[3]}/60 {round(mindset[3]*100/60,2)}%\n')
        f.write(f'Зеленый: {mindset[4]}/60 {round(mindset[4]*100/60,2)}%\n')
        f.write(f'Желтый: {mindset[5]}/60 {round(mindset[5]*100/60,2)}%\n')
        f.write(f'Бирюзовый: {mindset[6]}/60 {round(mindset[6]*100/60,2)}%\n\n')

        #f.write(f'Сумма баллов {sum(mindset)}')

        duress = get_duress_mindset(people)

        f.write("Мировоззрение во время стресса:\n\n")

        f.write(f'Фиолетовый: {duress[0]}/15 {round(duress[0]*100/15,2)}%\n')
        f.write(f'Красный: {duress[1]}/15 {round(duress[1]*100/15, 2)}%\n')
        f.write(f'Синий: {duress[2]}/15 {round(duress[2]*100/15,2)}%\n')
        f.write(f'Оранжевый: {duress[3]}/15 {round(duress[3]*100/15,2)}%\n')
        f.write(f'Зеленый: {duress[4]}/15 {round(duress[4]*100/15,2)}%\n')
        f.write(f'Желтый: {duress[5]}/15 {round(duress[5]*100/15,2)}%\n')
        f.write(f'Бирюзовый: {duress[6]}/15 {round(duress[6]*100/15,2)}%\n\n')

        #f.write(f'Сумма баллов {sum(duress)}')

        if sum(mindset) != 60 or sum(duress) != 15:
            print(filename, "Wrong test")

        states = get_change_state(people)

        f.write("Фазы изменений:\n\n")

        f.write(f"Альфа: {states[0]}/12\n")
        f.write(f"Бета: {states[1]}/12\n")
        f.write(f"Гамма: {states[2]}/12\n")
        f.write(f"Дельта: {states[3]}/12\n")
        f.write(f"Новая Альфа: {states[4]}/12\n\n")



    # with open(filename, encoding="utf-8") as f:
    #     data = f.read()

    #     send_email(people["info"][1], data)



    # data = people["answers"]

    # res = dict(zip(header["answers"], data))

    # buf = ""

    # for k,v in res.items():
    #     buf += f"{k}: {v}\n"

    # print(buf)

    # send_email(people["info"][1], buf)
