import speech_recognition as sr
import AppOpener as opener
import os
from googletrans import Translator
from pywhatkit import search
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
recog = sr.Recognizer()
micro = sr.Microphone()
_settings = {'voiceInput': True, 'textInput':True, 'textOutput':True, 'voiceOutput':True}
while True:
    with micro as source:
        recog.adjust_for_ambient_noise(source)
        audio = recog.listen(source)
        if _settings['voiceInput']:
            try:
                inp = recog.recognize_google(audio, language='ru-RU')
            except:
                continue
        if _settings['textInput']:
            inp = input()
        inp = inp.lower().split()
        print(inp)
        if 'выключи' in inp or 'выключить' in inp and 'компьютер' in inp or 'ноутбук' in inp:
            if _settings['textOutput']:
                print("Принято. Выключаю компьютер")
            if _settings['voiceOutput']:
                ...
            os.system("shutdown /s /t 1")
        if 'открой' in inp or 'запусти' in inp:
            try:
                open = inp.index('запусти')
            except:
                open = inp.index('открой')
            finally:
                try:
                    opener.open(inp[open+1],match_closest=True)
                    if _settings['textOutput']:
                        print("Принято. Запускаю приложение.")
                    if _settings['voiceOutput']:
                        ...
                except:
                    print("Такого приложения не существует")
        if 'выключи' in inp or 'закрой' in inp:
            try:
                open = inp.index('выключи')
            except:
                open = inp.index('закрой')
            finally:
                try:
                    if _settings['textOutput']:
                        print("Принято. Закрываю приложение.")
                    if _settings['voiceOutput']:
                        ...
                    opener.close(inp[open+1],match_closest=True)
                except:
                    print(f'Приложение {inp[open+1]} уже закрыто или не существует')
        if 'покажи' in inp and 'приложение' in inp or 'приложения' in inp:
            if _settings['textOutput']:
                print("Принято. Приложения, которые я могу открывать и закрывать:")
            if _settings['voiceOutput']:
                ...
            opener.open('ls')
        if 'спроси' in inp:
            if _settings['textOutput']:
                print("Вбиваю ваш вопрос в поиск.")
            if _settings['voiceOutput']:
                ...
            question = inp.index('спроси')
            search(" ".join(inp[question+1:]))
        elif 'что' in inp and 'такое' in inp:
            if _settings['textOutput']:
                print("Вбиваю ваш вопрос в поиск.")
            if _settings['voiceOutput']:
                ...
            question = inp.index('такое')
            search(" ".join(inp[question-1:]))
        elif 'почему' in inp:
            if _settings['textOutput']:
                print("Вбиваю ваш вопрос в поиск.")
            if _settings['voiceOutput']:
                ...
            question = inp.index('почему')
            search(" ".join(inp[question:]))
        if 'громкость' in inp:
            vol = inp.index('громкость')+1
            try:
                int(inp[vol])
            except:
                vol +=1
            try:
                volume.SetMasterVolumeLevel(int(-(-0.66*int(inp[vol])+66)), None)
                if _settings['textOutput']:
                    print(f"Ставлю громкость на {inp[vol]}%.")
                if _settings['voiceOutput']:
                    ...
            except:
                if _settings['textOutput']:
                    print("Вбиваю ваш ответ в поиск.")
                if _settings['voiceOutput']:
                    ...
        if 'переведи' in inp and 'английский' in inp:
            if _settings['textOutput']:
                print("Что конкретно перевести? Напишите/Скажите.")
            if _settings['voiceOutput']:
                ...
            if _settings['textInput']:
                text = input()
            elif _settings["voiceInput"]:
                audio = recog.listen(source)
                try:
                    text = recog.recognize_google(audio)
                except:
                    print("Ожидаю. Что нужно перевести? Напишите.")
                    text = input()
            if _settings['textOutput']:
                try:
                    print(f"Перевод: {Translator().translate(text,dest='ru-RU')}")
                except:
                    print("Не получилось перевести")
                    continue
            if _settings['voiceOutput']:
                ...






















        """if 'включи' in inp or 'открой' in inp:
            if 'chrome' in inp:
                print("Принято")
                try:
                    opener.open('google chrome', match_closest=True, output=True)
                except:
                    print('Никак')#сказать - не получилось
            if 'telegram' in inp:
                print("Открываю телеграм")
                opener.open('telegram', match_closest=True, output=False)
            if 'edge' in inp:
                print("Открываю microsoft edge")
                opener.open('microsoft edge', match_closest=True, output=False)
        if 'повысь громкость до' in inp or 'громкость' in inp or 'понизь громкость до' in inp:
            print('Принято')
        if 'покажи приложени' in inp:
            opener.open('ls')"""