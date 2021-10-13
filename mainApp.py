from genericpath import isfile
import PySimpleGUI as sg
from numpy import NaN
import speech_recognition as sr
import pandas as pd  # pip install numpy==1.19.3
import sounddevice as sd
import wavio as wv
import datetime
import openpyxl
import pygame
import random
import time
import os

from math import ceil
from pathlib import Path
from pydub import AudioSegment
from pydub.playback import play
from google.cloud import texttospeech_v1
from google.cloud import speech

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = r'toves-tts-8e3dc9b99ece.json'
settings = sg.UserSettings()
STTclient = speech.SpeechClient()
TTSclient = texttospeech_v1.TextToSpeechClient()
recognizer = sr.Recognizer()

ts = time.time()
st = datetime.datetime.fromtimestamp(
    ts).strftime('%d-%m-%Y_%H-%M-%S')

# outputDF = [
#     ['Input_ID', 'Input_Text', 'Pitch', 'Tone', 'Speed', 'Noise_Overlay',
#         'Gender', 'Output_Text', 'Output_Audio', 'Pass_or_Fail']
# ]
# currentRowList = []

outputDF = [
    ['Time_Stamp'],
    ['Input_ID'],
    ['Input_Text'],
    ['Pitch'],
    ['Tone'],
    ['Speed'],
    ['Noise_Overlay'],
    ['Gender'],
    ['Output_Text'],
    ['Output_Audio'],
    ['Pass_or_Fail']
]

currentExpStr = ''

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
               sg.Button('X', key='Exit', font='Any 10',  button_color=('white', 'red'))],
              [sg.Text('     OAU1COB',
                       background_color=DARK_HEADER_COLOR, font='Any 7', text_color=BORDER_COLOR)]
              ]

top = [[sg.Text('', size=(50, 1), justification='c', pad=BPAD_TOP, font='Any 20')],
       [sg.T(f'{i*25}-{i*34}') for i in range(7)], ]

block_3 = [[sg.Text('Completion percentage', font='Any 15', pad=(5, 15))],
           [sg.ProgressBar(100, orientation='h', size=(
               27, 20), key='progressbar', pad=(5, 15), bar_color=(DARK_HEADER_COLOR, BORDER_COLOR)), sg.Text('0%', key='progress-percent')],
           [sg.Button('Open Report', key='-OPENREPORT-', visible=False)]]


block_2 = [[sg.Text('Enter a filename:', pad=(5, 10), font='Any 15')],
           [sg.Input(sg.user_settings_get_entry(''),
                     key='-FILEIN-', pad=(5, 15), size=(50, 20)), sg.FileBrowse(key='-IN-', file_types=(('Excel Files', '*.xlsx'),))],
           [sg.B('Start Test', key='-START-', pad=(5, 20))]]

block_4 = [
    [sg.Output(size=(59, 19), key='-OUTPUT-', background_color='black', text_color='white')]]

layout = [[sg.Column(top_banner, size=(960, 60), pad=(0, 0), background_color=DARK_HEADER_COLOR)],
          [sg.Column([[sg.Column(block_2, size=(450, 150), pad=BPAD_LEFT_INSIDE)],
                      [sg.Column(block_3, size=(450, 150),  pad=BPAD_LEFT_INSIDE)]], pad=BPAD_LEFT, background_color=BORDER_COLOR),
           sg.Column(block_4, size=(450, 320), pad=BPAD_RIGHT)]]

window = sg.Window('Dashboard PySimpleGUI-Style', layout, margins=(0, 0),
                   background_color=BORDER_COLOR, no_titlebar=True, grab_anywhere=True)

outputStr = ''
i = 0


def googleSTT(fileName):
    filePath = r'STT_Inputs\\' + fileName

    with open(filePath, 'rb') as f1:
        byte_data_mp3 = f1.read()
    audio_mp3 = speech.RecognitionAudio(content=byte_data_mp3)

    config_mp3 = speech.RecognitionConfig(
        sample_rate_hertz=44400,
        enable_automatic_punctuation=True,
        language_code='en-US',
        audio_channel_count=2
    )

    response_standard_mp3 = STTclient.recognize(
        config=config_mp3,
        audio=audio_mp3
    )

    tempStr = ''

    for result in response_standard_mp3.results:
        windowUpdate('Transcript: {}'.format(
            result.alternatives[0].transcript))
        tempStr = tempStr + result.alternatives[0].transcript

    global outputDF
    global currentExpStr
    outputDF[8].append(tempStr)
    if(any(email_service in currentExpStr for email_service in tempStr)):
        outputDF[10].append('PASS')
    else:
        outputDF[10].append('FAIL')


def recordAudio(fileName, duration):
    frequency = 44400
    recording = sd.rec(int(duration * frequency),
                       samplerate=frequency, channels=2)
    sd.wait()

    wv.write(fileName, recording, frequency, sampwidth=2)


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


def googleTTS(inputNum, quote, outputCaptureTime):
    #global currentDFrow
    inputStr = str(inputNum) + '_' + st

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

    response = TTSclient.synthesize_speech(
        request={
            'input': input_text,
            'voice': voice,
            'audio_config': audio_config
        }
    )

    with open((r'TTS_Outputs\TTS_Output' + inputStr + '.mp3'), 'wb') as out:
        out.write(response.audio_content)

    windowUpdate('TTS_Output' + inputStr + '.mp3 created')
    file = Path('TTS_Outputs', 'TTS_Output' + inputStr + '.mp3')

    if(str(quote[4]) != 'nan' and str(quote[4]) != ''):
        # overlay noise
        sound1 = AudioSegment.from_file(
            Path('Noises', quote[4]))
        sound2 = AudioSegment.from_file(file)

        combined = sound2.overlay(sound1)
        combined.export(file, format='mp3')

    try:
        if Path(file).is_file():
            pygame.mixer.init()

            # play input
            pygame.mixer.music.load(file)
            pygame.mixer.music.play(0, 0.0)

            while pygame.mixer.music.get_busy() == True:
                continue

            windowUpdate('TTS_Output' + inputStr + '.mp3 played as input')

            # capture output for outputCaptureTime seconds
            outPath = str(Path('STT_Inputs', 'STT_Input' + inputStr + '.mp3'))
            recordAudio(outPath, outputCaptureTime)

            global currentExpStr
            global outputDF
            cwd = os.getcwd()

            outputDF[0].append(datetime.datetime.now())
            outputDF[1].append(quote[7])
            outputDF[2].append(quote[0])
            outputDF[3].append(quote[1])
            outputDF[4].append(quote[2])
            outputDF[5].append(quote[3])
            outputDF[6].append(quote[4])
            outputDF[7].append(quote[5])
            outputDF[9].append(Path(cwd, outPath))
            currentExpStr = quote[6]

    except Exception as e:
        windowUpdate('Error: ' + str(e) + ' ' + file)


def excelParse(data):

    inputList = data['MAIN']['Input_ID'].to_list()
    succeedingTimeList = data['MAIN']['Reply_Wait_Time'].to_list()
    tempList = zip(
        data['INPUT_CONSTRUCTOR']['Input_RawText'],
        data['INPUT_CONSTRUCTOR']['Pitch'],
        data['INPUT_CONSTRUCTOR']['Tone'],
        data['INPUT_CONSTRUCTOR']['Speed'],
        data['INPUT_CONSTRUCTOR']['Input_EnvironmentNoise'],
        data['INPUT_CONSTRUCTOR']['Gender'],
        data['INPUT_CONSTRUCTOR']['Output_Expected'],
        data['INPUT_CONSTRUCTOR']['Input_ID'],
    )
    dataDic = dict(zip(data['INPUT_CONSTRUCTOR']['Input_ID'],
                   tempList))

    for i in range(0, len(inputList)):
        googleTTS(i, dataDic[inputList[i]], succeedingTimeList[i])
        percent = ((i+1)*100)/(len(inputList))

        windowUpdate('TTS Audio Play: ' + dataDic[inputList[i]][0])
        # time.sleep(succeedingTimeList[i])
        window['progressbar'].update(ceil(percent))
        window['progress-percent'].update(str(ceil(percent)) + '%')


def processSTT():

    # To get only files present in a path
    fileList = os.listdir(path=r'STT_Inputs\\')

    for i in fileList:
        if(('.mp3' in i) and (st in i)):
            windowUpdate(i)
            googleSTT(i)


while True:
    # Event Loop
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == '-START-':
        filename = values['-IN-']
        if Path(filename).is_file():
            window['-START-'].update(disabled=True)
            try:
                df = pd.read_excel(filename, sheet_name=None)
                excelParse(df)
                windowUpdate(str(filename) + ' parsed')
                processSTT()

                # write output
                transsfDf = [list(i) for i in zip(*outputDF)]

                outputDFWrite = pd.DataFrame(transsfDf)
                outputDFWrite.to_excel(
                    'Output_Reports\\TOVES_Output_' + st + '.xlsx', sheet_name='MAIN', index=False, header=False)

            except Exception as e:
                windowUpdate('Error: ' + str(e))
            window['-START-'].update(disabled=False)
            window['-OPENREPORT-'].update(visible=True)
        else:
            windowUpdate('Please select an valid TC excel file')
    if event == '-OPENREPORT-':
        os.system('Output_Reports\\TOVES_Output_' + st + '.xlsx')
window.close()
