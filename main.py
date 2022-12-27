import pandas as pd
import random
import smtplib, ssl
from email.message import EmailMessage


def find_targets(df):
    assigned_people = []
    for person in df.index:
        available_people = set(df['name'])
        available_people -= set([df.loc[person]['name']])
        available_people -= set([df.loc[person]['partner']])
        available_people -= set(assigned_people)
        available_people = list(available_people)
        if len(available_people) < 1:
            assigned_people = find_targets(df)
            break
        else:
            target = random.choice(available_people)
            assigned_people.append(target)
    return assigned_people



if __name__ == '__main__':
    df = pd.read_excel('secret_santa_list.xlsx')
    ordered_targets = find_targets(df)
    df['target'] = ordered_targets

    port = 465  # For SSL
    email_acc = input('Type your google mail account and press enter: ')
    password = input("Type your password and press enter: ")

    # Create a secure SSL context
    context = ssl.create_default_context()

    for person in df.index:
        santa = df.loc[person]['name']
        receiver = df.loc[person]['target']
        msg = EmailMessage()
        msg.set_content('Hi {}!\n '
                        'Du beschenkst dieses Jahr {}\n '
                        'Bitte erinnere dich daran, dass das Budget ca. 30CHF ist.\n'
                        'Digital Santa wuenscht viel Freude beim schenken\n'
                        'Frohe Weihnachten!'.format(santa, receiver))

        msg['Subject'] = 'Secret Santa Familie Glettig'
        msg['From'] = email_acc
        msg['To'] = df.loc[person]['email']

        with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
            server.login(email_acc, password)
            server.send_message(msg)


# pagkylclxgcgjxrx





