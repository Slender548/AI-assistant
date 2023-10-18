import speech_recognition as sr
import AppOpener as opener
from pywhatkit import search
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
recog = sr.Recognizer()
micro = sr.Microphone()
while True:
    with micro as source:
        recog.adjust_for_ambient_noise(source)
        audio = recog.listen(source)
        #try:
        #    inp = recog.recognize_google(audio, language='ru-RU')
        #except:
        #    continue
        inp = input()
        inp = inp.lower().split()
        print(inp)
        if 'открой' in inp or 'запусти' in inp:
            try:
                open = inp.index('запусти')
            except:
                open = inp.index('открой')
            finally:
                try:
                    opener.open(inp[open+1],match_closest=True)
                except:
                    print("Такого приложения не существует")
        if 'выключи' in inp or 'закрой' in inp:
            try:
                open = inp.index('выключи')
            except:
                open = inp.index('закрой')
            finally:
                try:
                    opener.close(inp[open+1],match_closest=True)
                except:
                    print(f'Приложение {inp[open+1]} уже закрыто или не существует')
        if 'покажи' in inp and 'приложение' in inp or 'приложения' in inp:
            opener.open('ls')
        if 'спроси' in inp:
            question = inp.index('спроси')
            search(" ".join(inp[question+1:]))
        elif 'что' in inp and 'такое' in inp:
            question = inp.index('такое')
            search('что '+" ".join(inp[question:]))
        elif 'почему' in inp:
            question = inp.index('почему')
            search(" ".join(inp[question:]))
        if 'громкость до' in inp:
            vol = inp.index('до')
            if 'повысь' in inp or 'повысить' in inp:
                volume.SetMasterVolumeLevel()
            elif 'понизь' in inp or 'понизить' in inp:
                volume.SetMasterVolumeLevel()
            else:
                volume.SetMasterVolumeLevel()
        elif 'громкость' in inp:
            vol = inp.index('громкость')+1
            while int(inp[vol])=="":
                print(type(inp[vol]))
                if vol >= len(inp):
                    print('Какая громкость? Повторите')
                    continue
                vol += 1
            volume.SetMasterVolumeLevel(int(-(-0.66*int(inp[vol])+66)), None)
























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