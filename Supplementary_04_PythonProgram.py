# This script was tested on Python 3.10.8 in pyCharm 2023.1
# !!=== The openai.api_key must be filled with your own openAI API key before running, ===!!
# !!=== which can be applied at platform.openai.com.                                   ===!!
# Be sure you have installed essential packages, including openpyxl, pandas, openai, tqdm, etc.
# Input data file must follow a specific header specified in the InputTemplate.xlsx.
# Be sure the output file is not used by other applications, or you will lose all the output!
# A stable network connection is strongly recommended.
# Coding: Huaiyuan Ma
# Project created on May 15, 2023,
# Last update on Jun 24, 2023.
import json
import time
import pandas as pd
import openai
import math
import logging
from tqdm import tqdm

# openAI initialization.
openai.api_key = "openAI API key"

# openAI model selection

MODEL = "text-davinci-003"
# MODEL = "text-curie-001"

# prompt lines initialization. Feel free to change the prompt line as you want.

promptLine01 = '''Please tell me the pathological diagnosis, invasion depth, vertical margin status, horizontal status, vascular invasion status and lymphatic invasion status of the ESD specimen based on the following information in JSON format:
 { 
    "Pathological report": { 
        "Pathological diagnosis":"",
        "Invasion depth": "", 
        "Vertical margin":"", 
        "Horizontal margin":"", 
        "Vascular invasion":"", 
        "Lymphatic invasion":""
    } 
} 
Pathological diagnosis is the highest one. Non-cancerous lesion is recorded as Noncancerous, including chronic inflammation.
Invasion depth can be: 1. pT1a-EP, lesion within epithelial layer (M1); 2. pT1a-LPM, lesion invaded the lemina propria layer (M2); 3.pT1a-MM, lesion invaded the muscularis memberance (M3); 4. pT1b-SM, lesion invaded the submucosal layer.5. Deeper. High-graded dysplasia or low -graded dysplasia is categorized as pT1a-EP.  For inflammations,invasion depth is recorded as None.
Vertical margin involvement can be positive or negative. Vertical margin involvement is negtive when the distance between the dysplastic area and the oral, anal, anterior, and posterior margins are all measurable.
Horizontal margin involvement can be positive or negative. Horizontial margin is negtive when not involved explictly.
Vascular invasion status can be positive, or negative if not stated. 
Lymphatic invasion status can be positive, or negative if not stated. '''

promptLine02 = '''Your Prompt Here.'''
promptLine03 = '''Your Prompt Here.'''
promptLine04 = '''Your Prompt Here.'''
promptLine05 = '''Your Prompt Here.'''

# Choose the prompt you want.
promptLine = promptLine01

# Report initialization
PathoReport = ""

# input file provides pathological report you collected.
inputFile = r'Supplementary_04_InputTemplate.xlsx'

# This file stores the output.
outputFile = r'./GPT_Quality_Control.xlsx'
# This file stores the temp data.
tempFile = r'./tempOutput_ESD.xlsx'

logFile = r'./GPT_log.txt'

# logger init.
logging.basicConfig(filename=logFile, format='%(asctime)s %(message)s', filemode='w')
logger = logging.getLogger()
logger.setLevel(logging.INFO)

# read your dataset to dataframe from input file.
try:
    df_source = pd.read_excel(inputFile, header=0)
except FileNotFoundError as e:
    print(e)
    print('input file does not exist. Exit.')
    exit()

print("Start time:", time.asctime())

# Core algorithm for GPT job. Do not modify unless you know what you are doing.
for index, row in tqdm(df_source.iterrows()):
    print("Working on {0}th record of input file...".format(str(index)))
    logging.info("Working on {0}th record of input file.".format(str(index)))

    PathoReport = row['PathologicalReport']

    # deal with empty blanks occasionally appears in input file.
    try:
        if math.isnan(PathoReport):
            df_source.iat[index, df_source.columns.get_loc('ResponseData_GPT')] = 'NaN Detected.'
    except TypeError:
        # Deal with non-NaN data, normal data.
        # A simple way to deal with RateLimiter Exception.
        for attempt in range(5):
            if attempt == 4:
                print('retried for three times. Exit.')
                exit()

            # Do the translation job first.
            try:
                responseTranslation = openai.Completion.create(
                    engine=MODEL,
                    prompt="Translate into English:" + str.strip(PathoReport),
                    temperature=0,
                    max_tokens=256,
                    n=1,
                    stop=None,
                )
            except Exception as e:
                print(e, 'Sleep for 10s and retry.')
                time.sleep(10)
                continue

            # Submit the translation and prompt to GPT. A trick is used to avoid extra information ahead of JSON.
            try:
                response = openai.Completion.create(
                    engine=MODEL,
                    prompt=promptLine+'\n'+responseTranslation['choices'][0]['text']+'.\n\n',
                    temperature=0,
                    max_tokens=256,
                    n=1,
                    stop=None,
                )
            except Exception as e:
                print(e, 'Sleep for 10s and retry.')
                time.sleep(10)
                continue

            # Store response of GPT to dataframe.

            df_source.iat[index, df_source.columns.get_loc('Translation_GPT')] = \
                str.strip(responseTranslation['choices'][0]['text'])

            df_source.iat[index, df_source.columns.get_loc('ResponseData_GPT')] = \
                str.strip(response['choices'][0]['text'])

            logging.info('\n' + str(index) + '\n' + responseTranslation['choices'][0]['text'])
            logging.info('\n'+str(index)+'\n'+response['choices'][0]['text'])

            # Decode response date to each field.
            # In rare cases, the response JSON is ruptured and should be skipped.
            try:
                myjson_data = json.loads(response['choices'][0]['text'])

                df_source.iat[index, df_source.columns.get_loc('PathologicalDiagnosis_GPT')] = \
                    myjson_data['Pathological report']['Pathological diagnosis']

                df_source.iat[index, df_source.columns.get_loc('InvasionDepth_GPT')] = \
                    myjson_data['Pathological report']['Invasion depth']

                df_source.iat[index, df_source.columns.get_loc('VM_GPT')] = \
                    myjson_data['Pathological report']['Vertical margin']

                df_source.iat[index, df_source.columns.get_loc('HM_GPT')] = \
                    myjson_data['Pathological report']['Horizontal margin']

                df_source.iat[index, df_source.columns.get_loc('VI_GPT')] = \
                    myjson_data['Pathological report']['Vascular invasion']

                df_source.iat[index, df_source.columns.get_loc('LI_GPT')] = \
                    myjson_data['Pathological report']['Lymphatic invasion']

                # To avoid RateLimiter of openAI, sleep 1s for every request.
                time.sleep(1)
            except Exception as e:
                logging.info(str(index) + ' ' + str(e))
                break

            # In case of corruption, store output for every 10 requests in temp file.
            if int(str(index)) % 10 == 0:
                with pd.ExcelWriter(tempFile, mode='w', engine='openpyxl') as writer:
                    df_source.to_excel(writer, sheet_name="Sheet1")
                print("Output temperately stored.")

            # To avoid RateLimiter of openAI, sleep 10 secs for every 60 requests.
            if int(str(index)) % 60 == 0:
                print('Sleep for 10s to avoid RateLimiter of openAI.')
                time.sleep(10)
            break

print("End time:", time.asctime())

# Write the output to Excel file.
with pd.ExcelWriter(outputFile, mode='w', engine='openpyxl') as writer:
    df_source.to_excel(writer, sheet_name="Sheet1")
print("the output is in {0}.".format(outputFile))
