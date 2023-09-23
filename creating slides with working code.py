
import os
import google.auth
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.auth import credentials
from googleapiclient.errors import HttpError
import openai
import time
import requests
import json

"""
definition should include english 
and definition slide should include image
images from google search
google api for images
app that helps you create powerpoint slides built for dot AI
word multiple meaning
word as itle
type of word ie adj noun etc
highlight the word in each sentence
adding pinyin
put definition on the same slide as examples
Title Line: Target word, pinyin, English meaning
somplified and traditional
continue working on extension javascript (but not right now)
selecting all for words in title line for javascript extension
google ngram
"""

#openai.api_key = "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" # Replace with your OpenAI API key

SCOPES = ['https://www.googleapis.com/auth/presentations']

creds = None

file_path = "C:\\Users\\Asya\\OneDrive - Pomona College\\Desktop\\my_codes\\experiments\\venv\\Scripts\\token.json"

if os.path.exists('file_path'):
    creds = credentials.Credentials.from_authorized_user_file('file_path', SCOPES)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('webclient2.json', SCOPES)
        creds = flow.run_local_server(port=8080)

    with open(file_path, 'w') as token:
        token.write(creds.to_json())


def create_slide(service, presentation_id, word, definition=None):
    try:
        # Create a blank slide
        requests = [
            {
                'createSlide': {
                    'slideLayoutReference': {
                        'predefinedLayout': 'BLANK'
                    }
                }
            }
        ]

        body = {
            'requests': requests
        }

        response = service.presentations().batchUpdate(presentationId=presentation_id, body=body).execute()
        create_slide_response = response.get('replies')[0].get('createSlide')
        slide_id = create_slide_response.get('objectId')
        print(f"Created slide with ID: {slide_id}")

        # Create a textbox on the slide with a unique object ID
        textbox_object_id = f"{slide_id}_textbox"
        requests = [
            {
                'createShape': {
                    'objectId': textbox_object_id,
                    'shapeType': 'TEXT_BOX',
                    'elementProperties': {
                        'pageObjectId': slide_id,
                        'size': {
                            'height': {
                                'magnitude': 100,
                                'unit': 'PT'
                            },
                            'width': {
                                'magnitude': 300,
                                'unit': 'PT'
                            }
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 100,
                            'translateY': 100,
                            'unit': 'PT'
                        }
                    }
                }
            },
            {
                'insertText': {
                    'objectId': textbox_object_id,
                    'text': word
                }
            }
        ]

        body = {
            'requests': requests
        }

        response = service.presentations().batchUpdate(presentationId=presentation_id, body=body).execute()
        print("Text added to the slide.")

        # Append the definition to the existing textbox if user chooses to add the definition
        if definition:
            requests = [
                {
                    'insertText': {
                        'objectId': textbox_object_id,
                        'insertionIndex': len(word),  # Append to the beginning of the textbox
                        'text': definition
                    }
                }
            ]

            body = {
                'requests': requests
            }

            response = service.presentations().batchUpdate(presentationId=presentation_id, body=body).execute()
            print("Text added to the slide.")

    except HttpError as error:
        print(f"An error occurred: {error}")
        print("Slide not created")
        return error

    return response

def get_word_definition(word, type="simplified", api_key =  "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" ):
    url = "https://api.openai.com/v1/completions"
    model =  "text-davinci-003"
    prompt = f"Give me the definition of this chinese '{word}' in '{type}' chinese"
    max_tokens = 100
    temperature = 0.5
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    data = {
        "model": model,
        "prompt": prompt,
        "max_tokens": max_tokens,
        "temperature": temperature,
        "n": 1,
        "stop": None,
    }

    response = requests.post(url, headers=headers, json=data)
    response_data = response.json()

    choices = response_data["choices"][0]["text"].strip() if "choices" in response_data else ""

    return choices

def get_word_definition_lang(word, language, api_key =  "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" ):
    url = "https://api.openai.com/v1/completions"
    model =  "text-davinci-003"
    prompt = f"Give me the definition of this '{word}' in '{language}'"
    max_tokens = 100
    temperature = 0.5
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    data = {
        "model": model,
        "prompt": prompt,
        "max_tokens": max_tokens,
        "temperature": temperature,
        "n": 1,
        "stop": None,
    }

    response = requests.post(url, headers=headers, json=data)
    response_data = response.json()

    choices = response_data["choices"][0]["text"].strip() if "choices" in response_data else ""

    return choices


def three_ex(word, type = "simplified", level= "1-2", api_key =  "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" ):
    url = "https://api.openai.com/v1/completions"
    model =  "text-davinci-003"
    prompt = f"give me 3 sentences in '{type}' chinese of HSK level '{level}' using this word'{word}'"
    max_tokens = 100
    temperature = 0.5
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    data = {
        "model": model,
        "prompt": prompt,
        "max_tokens": max_tokens,
        "temperature": temperature,
        "n": 1,
        "stop": None,
    }

    response = requests.post(url, headers=headers, json=data)
    response_data = response.json()

    choices = response_data["choices"][0]["text"].strip() if "choices" in response_data else ""

    return choices


def three_ex_lang(word, language, level= "1-2", api_key =  "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" ):
    url = "https://api.openai.com/v1/completions"
    model =  "text-davinci-003"
    prompt = f"give me 3 sentences in '{language}' of level '{level}' using this word'{word}'"
    max_tokens = 100
    temperature = 0.5
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    data = {
        "model": model,
        "prompt": prompt,
        "max_tokens": max_tokens,
        "temperature": temperature,
        "n": 1,
        "stop": None,
    }

    response = requests.post(url, headers=headers, json=data)
    response_data = response.json()

    choices = response_data["choices"][0]["text"].strip() if "choices" in response_data else ""

    return choices


def create_presentation(title):
    service = build('slides', 'v1', credentials=creds)

    presentation = {
        'title': title,
        'pageSize': {
            'height': {'magnitude': 612, 'unit': 'PT'},
            'width': {'magnitude': 792, 'unit': 'PT'}
        },
        'slides': [
            {
                'slideProperties': {
                    'layoutObjectId': 'LAYOUT_ID_1'
                }
            }
        ]
    }

    try:
        presentation = service.presentations().create(body=presentation).execute()
        presentation_id = presentation['presentationId']
        return presentation_id
    except HttpError as error:
        print(f'An error occurred: {error}')

def create_slide_with_definition(service, presentation_id, word, type):
    definition = get_word_definition(word, type)
    if definition:
        response = create_slide(service, presentation_id, word, definition)
        if response is not None:
            print("Slide created successfully.")
    else:
        print("No definition available for the word.")


def create_slide_three_ex(service, presentation_id, word, type, level):
    definition = three_ex(word, type, level)
    if definition:
        response = create_slide(service, presentation_id, word, definition)
        if response is not None:
            print("Slide created successfully.")
    else:
        print("No examples available for the word.")


def create_slide_with_definition_lang(service, presentation_id, word, language):
    definition = get_word_definition_lang(word, language)
    if definition:
        response = create_slide(service, presentation_id, word, definition)
        if response is not None:
            print("Slide created successfully.")
    else:
        print("No definition available for the word.")

def create_slide_three_ex_lang(service, presentation_id, word, language, level):
    definition = three_ex_lang(word, language, level)
    if definition:
        response = create_slide(service, presentation_id, word, definition)
        if response is not None:
            print("Slide created successfully.")
    else:
        print("No examples available for the word.")

    

if __name__ == '__main__':
    title = input("what would you like to name this presentation?: ")
    presentation_id = create_presentation(title)
    print(f'Created presentation with ID: {presentation_id}')

    service = build('slides', 'v1', credentials=creds)

    add_slides = True

    while add_slides:
        user_input = input("Input words seperated by a space: "+ '\n'+ "if you're done making slides, type 'over' ")
        if user_input == 'over':
            add_slides = False
        level_user = input("Select dificulty level: 1-2 = HSK 1-2 2-3 = HSK 3-4 5-6 = HSK 5-6: ")
        type_user = input("Do you want the definition in simplified Chinese? Traditional Chinese? Both? (simplified/traditional/both simplified and traditional): ")
        list_of_words = user_input.split(" ")
        for word in list_of_words:
            create_slide_with_definition(service, presentation_id, word, type_user)
            create_slide_three_ex(service, presentation_id, word, type_user, level_user)

           


"""
def get_word_definition_simple(word, api_key =  "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" ):
    url = "https://api.openai.com/v1/completions"
    model =  "text-davinci-003"
    prompt = f"Give me the definition of this chinese'{word}'in simplified chinese"
    max_tokens = 100
    temperature = 0.5
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    data = {
        "model": model,
        "prompt": prompt,
        "max_tokens": max_tokens,
        "temperature": temperature,
        "n": 1,
        "stop": None,
    }

    response = requests.post(url, headers=headers, json=data)
    response_data = response.json()

    choices = response_data["choices"][0]["text"].strip() if "choices" in response_data else ""

    return choices

def get_word_definition_trad(word, api_key =  "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" ):
    url = "https://api.openai.com/v1/completions"
    model =  "text-davinci-003"
    prompt = f"'{word}'有什麼意思?"
    max_tokens = 100
    temperature = 0.5
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    data = {
        "model": model,
        "prompt": prompt,
        "max_tokens": max_tokens,
        "temperature": temperature,
        "n": 1,
        "stop": None,
    }

    response = requests.post(url, headers=headers, json=data)
    response_data = response.json()

    choices = response_data["choices"][0]["text"].strip() if "choices" in response_data else ""

    return choices
"""


"""
def three_ex_simp_int(word, api_key =  "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" ):
    url = "https://api.openai.com/v1/completions"
    model =  "text-davinci-003"
    prompt = f"给我三个HSK 3-4 级的句子用'{word}'"
    max_tokens = 100
    temperature = 0.5
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    data = {
        "model": model,
        "prompt": prompt,
        "max_tokens": max_tokens,
        "temperature": temperature,
        "n": 1,
        "stop": None,
    }

    response = requests.post(url, headers=headers, json=data)
    response_data = response.json()

    choices = response_data["choices"][0]["text"].strip() if "choices" in response_data else ""

    return choices

def three_ex_simp_adv(word, api_key =  "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" ):
    url = "https://api.openai.com/v1/completions"
    model =  "text-davinci-003"
    prompt = f"给我三个HSK 5-6 级的句子用'{word}'"
    max_tokens = 100
    temperature = 0.5
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    data = {
        "model": model,
        "prompt": prompt,
        "max_tokens": max_tokens,
        "temperature": temperature,
        "n": 1,
        "stop": None,
    }

    response = requests.post(url, headers=headers, json=data)
    response_data = response.json()

    choices = response_data["choices"][0]["text"].strip() if "choices" in response_data else ""

    return choices

def three_ex_trad_beg(word, api_key =  "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" ):
    url = "https://api.openai.com/v1/completions"
    model =  "text-davinci-003"
    prompt = f"給我三個HSK 1-2級的句子用'{word}'"
    max_tokens = 100
    temperature = 0.5
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    data = {
        "model": model,
        "prompt": prompt,
        "max_tokens": max_tokens,
        "temperature": temperature,
        "n": 1,
        "stop": None,
    }

    response = requests.post(url, headers=headers, json=data)
    response_data = response.json()

    choices = response_data["choices"][0]["text"].strip() if "choices" in response_data else ""

    return choices

def three_ex_trad_int(word, api_key =  "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" ):
    url = "https://api.openai.com/v1/completions"
    model =  "text-davinci-003"
    prompt = f"給我三個HSK 3-4級的句子用'{word}'"
    max_tokens = 100
    temperature = 0.5
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    data = {
        "model": model,
        "prompt": prompt,
        "max_tokens": max_tokens,
        "temperature": temperature,
        "n": 1,
        "stop": None,
    }

    response = requests.post(url, headers=headers, json=data)
    response_data = response.json()

    choices = response_data["choices"][0]["text"].strip() if "choices" in response_data else ""

    return choices

def three_ex_trad_adv(word, api_key =  "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" ):
    url = "https://api.openai.com/v1/completions"
    model =  "text-davinci-003"
    prompt = f"給我三個HSK 5-6級的句子用'{word}'"
    max_tokens = 100
    temperature = 0.5
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    data = {
        "model": model,
        "prompt": prompt,
        "max_tokens": max_tokens,
        "temperature": temperature,
        "n": 1,
        "stop": None,
    }

    response = requests.post(url, headers=headers, json=data)
    response_data = response.json()

    choices = response_data["choices"][0]["text"].strip() if "choices" in response_data else ""

    return choices
"""        
           
"""
def create_slide_with_definition_simple(service, presentation_id, word):
    definition = get_word_definition_simple(word)
    if definition:
        response = create_slide(service, presentation_id, word, definition)
        if response is not None:
            print("Slide created successfully.")
    else:
        print("No definition available for the word.")
  

def create_slide_with_definition_trad(service, presentation_id, word):
    definition = get_word_definition_trad(word)
    if definition:
        response = create_slide(service, presentation_id, word, definition)
        if response is not None:
            print("Slide created successfully.")
    else:
        print("No definition available for the word.")
   
"""        

"""
def create_slide_three_ex_beg_simp(service, presentation_id, word):
    definition = three_ex_simp_beg(word)
    if definition:
        response = create_slide(service, presentation_id, word, definition)
        if response is not None:
            print("Slide created successfully.")
    else:
        print("No examples available for the word.")

def create_slide_three_ex_int_simp(service, presentation_id, word):
    definition = three_ex_simp_int(word)
    if definition:
        response = create_slide(service, presentation_id, word, definition)
        if response is not None:
            print("Slide created successfully.")
    else:
        print("No examples available for the word.")

def create_slide_three_ex_adv_simp(service, presentation_id, word):
    definition = three_ex_simp_adv(word)
    if definition:
        response = create_slide(service, presentation_id, word, definition)
        if response is not None:
            print("Slide created successfully.")
    else:
        print("No examples available for the word.")

def create_slide_three_ex_beg_trad(service, presentation_id, word):
    definition = three_ex_trad_beg(word)
    if definition:
        response = create_slide(service, presentation_id, word, definition)
        if response is not None:
            print("Slide created successfully.")
    else:
        print("No examples available for the word.")

def create_slide_three_ex_int_trad(service, presentation_id, word):
    definition = three_ex_trad_int(word)
    if definition:
        response = create_slide(service, presentation_id, word, definition)
        if response is not None:
            print("Slide created successfully.")
    else:
        print("No examples available for the word.")

def create_slide_three_ex_adv_trad(service, presentation_id, word):
    definition = three_ex_trad_adv(word)
    if definition:
        response = create_slide(service, presentation_id, word, definition)
        if response is not None:
            print("Slide created successfully.")
    else:
        print("No examples available for the word.")
"""

                
"""
        if user_input.lower() == 'yes':
            word = input("Enter the word for the slide: ")
            get_definition_input = input("Do you want to get the definition? (yes/no): ")
            get_definition = get_definition_input.lower() == 'yes'

            if get_definition:
                simplified_input = input("Do you want the definition in simplified Chinese? (yes/no/both): ")
                simplified_y = simplified_input.lower() == 'yes'
                simplified_n = simplified_input.lower() == 'no'
                simplified_b = simplified_input.lower() == 'both'

                if simplified_y:
                    definition = get_word_definition_simple(word)
                    create_slide_with_definition_simple(service, presentation_id, word)
                elif simplified_n:
                    definition = get_word_definition_trad(word)
                    create_slide_with_definition_trad(service, presentation_id, word)
                else:
                    definition_s = get_word_definition_simple(word)
                    create_slide_with_definition_simple(service, presentation_id, word)
                    definition_t = get_word_definition_trad(word)
                    create_slide_with_definition_trad(service, presentation_id, word)


            get_example_input = input("Do you want to get three examples of the word in use? (yes/no): ")
            get_example = get_example_input.lower() == 'yes'

            if get_example:
                simplified_input = input("Do you want the definition in simplified Chinese? (yes/no/both): ")
                simplified_y = simplified_input.lower() == 'yes'
                simplified_b = simplified_input.lower() == 'both'

                if simplified_y:
                    get_example_choice_s = input("Select dificulty level: beg = HSK 1-2 int = HSK 3-4 adv = HSK 5-6: ")
                    example_choice_beg_s = get_example_choice_s.lower() == 'beg'
                    example_choice_int_s = get_example_choice_s.lower() == 'int'
                    example_choice_adv_s = get_example_choice_s.lower() == 'adv'

                    if example_choice_beg_s:
                        example_beg = three_ex_simp_beg(word)
                        create_slide_three_ex_beg_simp(service, presentation_id, word)
                    elif example_choice_int_s:
                        example_int = three_ex_simp_int(word)
                        create_slide_three_ex_int_simp(service, presentation_id, word)
                    else:
                        example_adv = three_ex_simp_adv(word)
                        create_slide_three_ex_adv_simp(service, presentation_id, word)
                
                elif simplified_b:
                    get_example_choice = input("Select dificulty level: beg = HSK 1-2 int = HSK 3-4 adv = HSK 5-6: ")
                    example_choice_beg = get_example_choice.lower() == 'beg'
                    example_choice_int = get_example_choice.lower() == 'int'
                    example_choice_adv = get_example_choice.lower() == 'adv'

                    if example_choice_beg:
                        example_beg_s = three_ex_simp_beg(word)
                        create_slide_three_ex_beg_simp(service, presentation_id, word)
                        example_beg_t = three_ex_trad_beg(word)
                        create_slide_three_ex_beg_trad(service, presentation_id, word)
                    elif example_choice_int:
                        example_int_s = three_ex_simp_int(word)
                        create_slide_three_ex_int_simp(service, presentation_id, word)
                        example_int_t = three_ex_trad_int(word)
                        create_slide_three_ex_int_trad(service, presentation_id, word)
                    else:
                        example_adv_s = three_ex_simp_adv(word)
                        create_slide_three_ex_adv_simp(service, presentation_id, word)
                        example_adv_t = three_ex_trad_adv(word)
                        create_slide_three_ex_adv_trad(service, presentation_id, word)


                else:
                    get_example_choice_t = input("Select dificulty level: beg = HSK 1-2 int = HSK 3-4 adv = HSK 5-6: ")
                    example_choice_beg_t = get_example_choice_t.lower() == 'beg'
                    example_choice_int_t = get_example_choice_t.lower() == 'int'
                    example_choice_adv_t = get_example_choice_t.lower() == 'adv'

                    if example_choice_beg_t:
                        example_beg = three_ex_trad_beg(word)
                        create_slide_three_ex_beg_trad(service, presentation_id, word)
                    elif example_choice_int_t:
                        example_int = three_ex_trad_int(word)
                        create_slide_three_ex_int_trad(service, presentation_id, word)
                    else:
                        example_adv = three_ex_trad_adv(word)
                        create_slide_three_ex_adv_trad(service, presentation_id, word)

        else:
            add_slides = False
        """