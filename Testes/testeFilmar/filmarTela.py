import cv2
import keyboard
import pyautogui
import numpy as np

fps = 30 #define quantos prints serão tirados por segundo da tela
tamanho_tela = tuple(pyautogui.size()) #define a resolução do vídeo baseado na própria tela

i = 1

nomeVideo = f'video{i}.avi' ## Criar uma forma de mudar o nome do vídeo automaticamente
i+=1
codec = cv2.VideoWriter_fourcc(*"XVID") #codificador do vídeo AVI
video = cv2.VideoWriter(nomeVideo, codec, fps, tamanho_tela) #nome do vídeo, codificador, taxa de quadros e resolução

while True:
    frame = pyautogui.screenshot() #tira 30 prints por segundo da tela
    frame = np.array(frame) #armazena os prints em um array
    
    frame = cv2.cvtColor(frame, cv2.COLOR_RGB2BGR) #Define as cores para o padrão do cv2
    
    video.write(frame) #junta as imagens em um vídeo
    
    if keyboard.is_pressed("esc"):
        break
    

video.release() #cria o vídeo
cv2.destroyAllWindows()