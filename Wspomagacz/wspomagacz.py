import re
import xlrd as xlrd
import subprocess as sp


def find_most_fitting_standard_char(compared_letter, my_dict):

    ranking_of_chars = {}
    min_difference = 1  # zaporowa wartosc, nigdy nie zostanie wybrana - nwm jak zrobic nieinitializowana zmienna
    for dict_element in my_dict:
        current_difference = lettersFrequency[compared_letter] - standardLetterFrequency[dict_element]
        current_difference_absolute = abs(current_difference)
        ranking_of_chars[dict_element] = current_difference_absolute
        if current_difference_absolute < min_difference:
            min_difference = current_difference_absolute
    return ranking_of_chars


def get_key(my_dict, val):
    for key, value in my_dict.items():
        if val == value:
            return key


alphabet = "abcdefghijklmnopqrstuvwxyz"  # litery alfabetu
dictAlphabet = {}

encryptedFile = open("kryptogram.txt", "r")
encryptedText = encryptedFile.read()
decryptedText = str(encryptedText)

#liczymy sumy wystepowan liter w tekscie zaszyfrowanym

occurences = []
sumOfUsages = 0
for index in alphabet:
    dictAlphabet[index] = encryptedText.count(index)
    occurences.append(dictAlphabet[index])
    sumOfUsages += dictAlphabet[index]

#liczymy czestotliwosc wystepowania liter w tekscie zaszyfrowanym

lettersFrequency = {}
for char in dictAlphabet:
    popularity = dictAlphabet[char] / sumOfUsages
    lettersFrequency[char] = popularity

# czestotliwosc wystepowania liter w jezyku polskim, skopiowane z wiki, to samo bylo podawne na wykladzie
# polskie znaki dodane do odpowiadajacych im "standardowych" liter, czestotliwosc o = o + ó

standardLetterFrequency = {}
czestotliwosc_wzorcowa = xlrd.open_workbook('CzestotliwoscWzorcowa.xls')
arkusz = czestotliwosc_wzorcowa.sheet_by_index(0)
for i in range(26):
    cell_value_class = arkusz.cell(i, 0).value
    cell_value_id = arkusz.cell(i, 1).value
    standardLetterFrequency[cell_value_class] = cell_value_id

# decryptingDict = litera w szyfrze: litera w tekscie jawnym
# najpierw definiuje symbole ktore nie sa literami i w kryptogramie sa takie same jak w tekscie po odszyfrowaniu

decryptingDict = {" ": " ", "\n": "\n", "*": "*"}

# posluzmy sie tez zasadami pisowni jezyka polskiego. Jakie litery moga wystepowac same?

singleLetters = ['a', 'i', 'o', 'u', 'w', 'z']
standardLetterFrequencySingleLetters = {}
for char in singleLetters:
    standardLetterFrequencySingleLetters[char] = standardLetterFrequency[char]

# teraz sprawdzam jakie litery wystepuja same w kryptogramie

singleLettersEncrypted = set(re.findall("[ ].[ ]", encryptedText))
singleLettersEncryptedList = []
for element in singleLettersEncrypted:
    element = element.strip()
    if element != "":
        singleLettersEncryptedList.append(element)

# sprawdzamy ktora litery w alfabecie i kryptogramie maja najblizsze siebie czestotliwosci wystepowania

for letter in lettersFrequency:

    if letter in singleLettersEncryptedList:
        rankingOfFittingChars = find_most_fitting_standard_char(letter, singleLetters)
        sorted_values = sorted(rankingOfFittingChars.values())
        matching_letter = get_key(rankingOfFittingChars, sorted_values[0])

    else:
        rankingOfFittingChars = find_most_fitting_standard_char(letter, standardLetterFrequency)
        sorted_values = sorted(rankingOfFittingChars.values())
        matching_letter = get_key(rankingOfFittingChars, sorted_values[0])

    if matching_letter not in decryptingDict.values():
        decryptingDict[letter] = matching_letter

# niektore litery alfabetu sa najbardziej pasujace dla kilku liter szyfru, dlatego teraz wybierzemy dla nich
# drugie lub dalsze dopasowania, tak by kazda litera szyfru miala unikalny odpowiednik

for char in dictAlphabet.keys():
    index = 1
    if char not in decryptingDict.keys():
        rankingOfFittingChars = find_most_fitting_standard_char(char, standardLetterFrequency)
        sorted_values = sorted(rankingOfFittingChars.values())
        matching_letter = get_key(rankingOfFittingChars, sorted_values[index])
        while matching_letter in decryptingDict.values():
            index += 1
            matching_letter = get_key(rankingOfFittingChars, sorted_values[index])
        decryptingDict[char] = matching_letter

# Teraz python pokaze nam co wyanalizowal. Dzieki jego pomocy powinnismy juz widziec jak rozszyfrowac reszte tekstu
# za pomoca podmiany jednej litery na inna. Jesli w "odszyfrowanym" przez komputer tekscie na pierwszy rzut oka
# widac blad, mowimy "zamiana" i pokazujemy, co ma zamienic. Np. jesli od razu widze, ze w danym miejscu powinno byc
# "b", a jest "c", mowie "zamiana", a potem "b" i "c" (lub "c"i "b", wszystko jedno)

isDecrypted = False

while isDecrypted is False:

    decryptedFile = open("tekstJawny.txt", "w")

    for char in encryptedText:
        if char in decryptingDict:
            decryptedFile.write(decryptingDict[char])

    programName = "notepad.exe"
    fileName = "tekstJawny.txt"
    sp.Popen([programName, fileName])

    isReplyCorrect = False  # walidacja inputu
    while isReplyCorrect is False:
        print("Wpisz OK by zaakceptować tekst lub ZAMIANA by zamienić 2 litery miejscami")
        userReply = input()
        userReplyUppercase = userReply.upper()
        if userReplyUppercase == "OK":
            isReplyCorrect = True
            isDecrypted = True
        elif userReplyUppercase == "ZAMIANA":
            isReplyCorrect = True
            print("Wpisz pierwszą literę:")
            firstLetter = input()
            print("Wpisz drugą literę:")
            secondLetter = input()
            # zamiana przypisania liter alfabetu do liter w szyfrie w slowniku
            firstLetterKey = get_key(decryptingDict, firstLetter)
            secondLetterKey = get_key(decryptingDict, secondLetter)
            helper = decryptingDict[firstLetterKey]
            decryptingDict[firstLetterKey] = decryptingDict[secondLetterKey]
            decryptingDict[secondLetterKey] = helper
        else:
            print("Błędna odpowiedź")
            isReplyCorrect = False
