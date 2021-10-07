import PySimpleGUI as sg
import pandas as pd  # pip install numpy==1.19.3
import datetime
import pygame
import pydub
import time
import os

from pathlib import Path
from pydub import AudioSegment
from pydub.playback import play
from google.cloud.texttospeech_v1.types.cloud_tts import SsmlVoiceGender
from google.cloud import texttospeech  # outdated or incomplete comparing to v1
from google.cloud import texttospeech_v1

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = r"toves-tts-8e3dc9b99ece.json"
settings = sg.UserSettings()

theme_dict = {'BACKGROUND': '#2B475D',
              'TEXT': '#FFFFFF',
              'INPUT': '#F2EFE8',
              'TEXT_INPUT': '#000000',
              'SCROLL': '#F2EFE8',
              'BUTTON': ('#000000', '#C2D4D8'),
              'PROGRESS': ('#FFFFFF', '#C7D5E0'),
              'BORDER': 1, 'SLIDER_DEPTH': 0, 'PROGRESS_DEPTH': 0}

# sg.theme_add_new('Dashboard', theme_dict)     # if using 4.20.0.1+
sg.LOOK_AND_FEEL_TABLE['Dashboard'] = theme_dict
sg.theme('Dashboard')

BORDER_COLOR = '#C7D5E0'
DARK_HEADER_COLOR = '#1B2838'
BPAD_TOP = ((20, 20), (20, 10))
BPAD_LEFT = ((20, 10), (0, 10))
BPAD_LEFT_INSIDE = (0, 10)
BPAD_RIGHT = ((10, 20), (10, 20))

top_banner = [[sg.Text('TOVES Demo' + ' '*89, pad=(15, 0), font='Any 20', background_color=DARK_HEADER_COLOR),
               sg.Button('X', key='Exit', font='Any 10',  button_color=('white', 'red'))]]

top = [[sg.Text('The Weather Will Go Here', size=(50, 1), justification='c', pad=BPAD_TOP, font='Any 20')],
       [sg.T(f'{i*25}-{i*34}') for i in range(7)], ]

block_3 = [[sg.Text('Completion percentage', font='Any 15', pad=(5, 15))],
           [sg.ProgressBar(100, orientation='h', size=(
               27, 20), key='progressbar', pad=(5, 15), bar_color=(DARK_HEADER_COLOR, BORDER_COLOR)), sg.Text('0%', key='progress-percent')]]


block_2 = [[sg.Text('Enter a filename:', pad=(5, 10), font='Any 15')],
           [sg.Input(sg.user_settings_get_entry(''),
                     key='-FILEIN-', pad=(5, 15), size=(50, 20)), sg.FileBrowse(key="-IN-")],
           [sg.B('Start Test', key='-START-', pad=(5, 20))]]

block_4 = [
    [sg.Output(size=(59, 19), key='-OUTPUT-', background_color='black', text_color='white')]]

layout = [[sg.Column(top_banner, size=(960, 60), pad=(0, 0), background_color=DARK_HEADER_COLOR)],
          [sg.Column(top, size=(920, 90), pad=BPAD_TOP)],
          [sg.Column([[sg.Column(block_2, size=(450, 150), pad=BPAD_LEFT_INSIDE)],
                      [sg.Column(block_3, size=(450, 150),  pad=BPAD_LEFT_INSIDE)]], pad=BPAD_LEFT, background_color=BORDER_COLOR),
           sg.Column(block_4, size=(450, 320), pad=BPAD_RIGHT)]]

window = sg.Window('Dashboard PySimpleGUI-Style', layout, margins=(0, 0),
                   background_color=BORDER_COLOR, no_titlebar=True, grab_anywhere=True)

outputStr = ''
i = 0


def translate(value, leftMin, leftMax, rightMin, rightMax):
    leftSpan = leftMax - leftMin
    rightSpan = rightMax - rightMin

    valueScaled = float(value - leftMin) / float(leftSpan)
    return rightMin + (valueScaled * rightSpan)


def windowUpdate(text):
    global outputStr
    ts = time.time()
    st = datetime.datetime.fromtimestamp(ts).strftime('%H:%M:%S.%f')
    outputStr = '\n' + st + ' : ' + text + '\n' + outputStr
    window['-OUTPUT-'].update(outputStr)
    # time.sleep(2)


def googleTTS(inputNum, quote):

    inputStr = str(inputNum)

    client = texttospeech_v1.TextToSpeechClient()

    input_text = texttospeech_v1.SynthesisInput(text=quote[0])

    nameVoice = 'en-IN-Wavenet-B'
    if (quote[5] == 'F'):
        nameVoice = 'en-IN-Wavenet-A'

    speakingRate = translate(quote[3], 1, 100, 0.25, 4)
    pitchNum = translate(quote[1], 1, 100, -20, 20)
    volume = translate(quote[2], 1, 100, -96, 16)
    # pitch ranges from -20 to 20 | volume ranges from -96 to 16 | speaking rate ranges from 0.25 to 4

    windowUpdate(str(speakingRate) + ' - ' +
                 str(pitchNum) + ' - ' + str(volume))

    voice = texttospeech_v1.VoiceSelectionParams(
        language_code='en-IN',
        name=nameVoice
    )

    audio_config = texttospeech_v1.AudioConfig(
        audio_encoding=texttospeech_v1.AudioEncoding.MP3,
        speaking_rate=speakingRate,
        pitch=pitchNum,
        volume_gain_db=volume
    )

    response = client.synthesize_speech(
        request={
            "input": input_text,
            "voice": voice,
            "audio_config": audio_config
        }
    )

    with open((r'TTS_Outputs\TTS_Output' + inputStr + '.mp3'), "wb") as out:
        out.write(response.audio_content)

    windowUpdate('TTS_Output' + inputStr + '.mp3 created')

    file = (r'TTS_Outputs\TTS_Output' + inputStr + '.mp3')

    try:
        if Path(file).is_file():
            pygame.mixer.init()
            # mixer.music.load(file)

            pygame.mixer.music.load(file)
            pygame.mixer.music.play(0, 0.0)
            while pygame.mixer.music.get_busy() == True:
                continue

            windowUpdate('TTS_Output' + inputStr + '.mp3 played as input')

    except Exception as e:
        windowUpdate('Error: ' + str(e) + ' ' + file)


def excelParse(data):

    inputList = data['MAIN']['Input_ID'].to_list()
    succeedingTimeList = data['MAIN']['Succeeding_Wait_Time'].to_list()
    tempList = zip(
        data['INPUT_CONSTRUCTOR']['Input_RawText'],
        data['INPUT_CONSTRUCTOR']['Pitch'],
        data['INPUT_CONSTRUCTOR']['Tone'],
        data['INPUT_CONSTRUCTOR']['Speed'],
        data['INPUT_CONSTRUCTOR']['Input_EnvironmentNoise'],
        data['INPUT_CONSTRUCTOR']['Gender'],
        data['INPUT_CONSTRUCTOR']['Output_Expected']
    )
    dataDic = dict(zip(data['INPUT_CONSTRUCTOR']['Input_ID'],
                   tempList))

    for i in range(0, len(inputList)):
        googleTTS(i, dataDic[inputList[i]])

        percent = ((i+1)*100)/(len(inputList))

        windowUpdate('TTS Audio Play: ' + dataDic[inputList[i]][0])
        time.sleep(succeedingTimeList[i])
        window['progressbar'].update(percent)
        window['progress-percent'].update(str(percent) + '%')


while True:             # Event Loop
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == '-START-':
        filename = values['-IN-']
        if Path(filename).is_file():
            try:
                df = pd.read_excel(filename, sheet_name=None)
                excelParse(df)
                windowUpdate(str(filename) + ' parsed')
            except Exception as e:
                windowUpdate('Error: ' + str(e))

window.close()
