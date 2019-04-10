import pandas as pd
import Person


def matchOne(one, two):
    return (
        one.location == two.location and
        one.capability == two.capability and
        one.market == two.market)


def matchTwo(one, two):
    return (
        one.location == two.location and
        two.capability == two.capability
    )


def matchThree(one, two):
    return one.location == two.location


# Read excel file to dataframe
def readMentorList():
    df = pd.read_excel("mentor-data.xlsx", skipinitialspace=True, dtype=str)

    # Strip whitespace & replace 'nan' values
    df = df.replace(' ', '')
    df = df.replace('nan', '')

    # Converts dataframe to dictionary list
    mentor_list = df.to_dict(orient='records')

    new_mentor_list = []

    for mentor in mentor_list:
        first = mentor['First']
        last = mentor['Last']
        location = mentor['Location']
        capability = mentor['Capability']
        market = mentor['Market']
        email = ''
        # Instantiate new Person
        person = Person.Person(first, last, location, capability, market, email, True)
        new_mentor_list.append(person)
    return new_mentor_list


# Read excel file to dataframe
def readMenteeList():
    df = pd.read_excel("mentee-data.xlsx", skipinitialspace=True, dtype=str)

    # Strip whitespace & replace 'nan' values
    df = df.replace(' ', '')
    df = df.replace('nan', '')

    # Converts dataframe to dictionary list
    mentee_list = df.to_dict(orient='records')

    new_mentee_list = []

    for mentee in mentee_list:
        first = mentee['First'].strip()
        last = mentee['Last'].strip()
        location = mentee['Location'].strip()
        capability = mentee['Capability'].strip()
        market = mentee['Market'].strip()
        email = ''
        # Instantiate new Person
        person = Person.Person(first, last, location, capability, market, email, False)
        new_mentee_list.append(person)
    return new_mentee_list


def toMatchDict(mentor, mentee):
    dict = {
        'Mentor': mentor.full_name,
        'Mentee': mentee.full_name,
        'Mentee Location': mentee.location
    }
    return dict


def outputToExcel(match_list):
    new_list = []
    for match in match_list:
        match_dictionary = toMatchDict(match['Mentor'], match['Mentee'])
        new_list.append(match_dictionary)
    match_list = new_list
    print(match_list)

    try:
        df = pd.DataFrame(match_list)
        writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter', options={'strings_to_urls': False})
        df.to_excel(writer)
        writer.save()
    except Exception as e:
        print(e)


def main():
    # Setup mentor and mentee object lists
    mentor_list = readMentorList()
    mentee_list = readMenteeList()

    match_list = []

    # Do match One
    for mentor in mentor_list:
        for mentee in mentee_list:
            if matchOne(mentor, mentee):
                match = {
                    'Mentor': mentor,
                    'Mentee': mentee
                }
                match_list.append(match)
                mentor_list.remove(mentor)
                mentee_list.remove(mentee)
                # break
                break

    # Do match Two
    for mentor in mentor_list:
        for mentee in mentee_list:
            if matchTwo(mentor, mentee):
                match = {
                    'Mentor': mentor,
                    'Mentee': mentee
                }
                match_list.append(match)
                mentor_list.remove(mentor)
                mentee_list.remove(mentee)
                break

    # Do match Three
    for mentor in mentor_list:
        for mentee in mentee_list:
            if matchThree(mentor, mentee):
                match = {
                    'Mentor': mentor,
                    'Mentee': mentee
                }
                match_list.append(match)
                mentor_list.remove(mentor)
                mentee_list.remove(mentee)
                break

    # Match leftovers
    for mentor in mentor_list:
        for mentee in mentee_list:
            match = {
                'Mentor': mentor,
                'Mentee': mentee
            }
            match_list.append(match)
            mentee_list.remove(mentee)
            break

    print('Total Matches:', len(match_list))
    outputToExcel(match_list)


# Instantiate main 888888888888888888888888888888888888888888888888888888888888
if __name__ == "__main__":
    main()
