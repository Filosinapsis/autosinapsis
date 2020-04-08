# -*- coding: utf-8 -*-
import json
import openpyxl
import codecs
from pathlib import Path

filosinapsis_file = Path('Red_Dependencia_temas.xlsx')
workbook = openpyxl.load_workbook(filosinapsis_file)
sheet = workbook.active

alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']


neuron_youtube_links = {}
neuron_id = 0
have_parent_neuron = False
topics = {}

for number in range(4, 301):
    for letter_index, letter in enumerate(alphabet):
        cell = letter + str(number)
        if sheet[cell].value == '' or sheet[cell].value == None:
            pass
        else:
            current_neuron = sheet[cell].value
            if neuron_id == 0:

                have_parent_neuron = False
                neuron_key_id = str(neuron_id)

                neuron_name = sheet[cell].value
                neuron_sub = [1]
                neuron_super = -1
                color = 'white'
                size = 'big'

                # Obteniendo datos...
                topics[neuron_key_id] = {
                    'id': neuron_id,
                    'name': neuron_name,
                    'SubNeurons':neuron_sub,
                    'SuperNeuron': neuron_super,
                    'color': color,
                    'size': size
                }
                neuron_id += 1
            else:
                have_parent_neuron = True
                travel_letter = alphabet[letter_index - 1]
                travel_number = number - 1
                possible_cell = travel_letter + str(travel_number)
                hypothesis_neuron = sheet[possible_cell].value
                
                row_back = 1
                if hypothesis_neuron == None or hypothesis_neuron == '' or hypothesis_neuron.startswith('1:'):
                    
                    parent_neuron_not_founded = True

                    while parent_neuron_not_founded:
                        
                        travel_letter = alphabet[letter_index - 1]
                        travel_number = number - row_back
                        possible_cell = travel_letter + str(travel_number)
                        hypothesis_neuron = sheet[possible_cell].value

                        if hypothesis_neuron == None or hypothesis_neuron == '':
                            pass
                        else:
                            print('El tema => ', end='')
                            print(hypothesis_neuron, end=' --> ')
                            print(f'contiene... => {current_neuron}')
                            parent_neuron_not_founded = False

                        row_back += 1

                else:
                    print('El tema => ', end='')
                    print(hypothesis_neuron, end=' --> ')
                    print(f'contiene... => {current_neuron}')

                    # Obteniendo datos...





print("Algoritmo Filosinapsis")
print("--> RESULTADO <--")
json_topics = json.dumps(topics, indent=4)
print(json_topics)

