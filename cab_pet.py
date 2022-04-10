#Script to generate first paragraphs in briefs


import pyperclip

def complete_brief(name, position, expose=True):
    if expose:
        pyperclip.copy(f'{name}, já qualificado nos autos do processo em referência, no qual figura como {position} vem, por seu/uas advogado/as, expor para ao final requerer o que segue.')
    else:
        pyperclip.copy(f'{name}, já qualificado nos autos do processo em referência, no qual figura como {position} vem, por seu/uas advogado/as, apresentar XX, nos termos seguintes.')


name = input('Qual o nome da parte (em letras maiúsculas)? ')
position = input(f'Qual a posição processual de {name}? ')
pet_inom = input('Petição inominada? ')
if  pet_inom == 'sim' or pet_inom == 'Sim':
     expose = True 
else : expose = False

complete_brief(name, position, expose=expose)

