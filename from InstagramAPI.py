import time
import webbrowser
import pyautogui
import PySimpleGUI as sg
import pandas as pd 

sg.theme('DarkBlue4')
layout = [
    [sg.Text("ESCOLHA UM INCIDENTE:")],
    [sg.Button('TIMEOUTS PARA ADQUIRENTES', size =(15, 2))], 
    [sg.Button('TIMEOUTS INICIO/FIM', size =(16, 2))],
    [sg.Button('TIMEOUTS PARA ADQUIRENTES SITEF', size =(16, 2))],
    [sg.Button('ADQUIRENTE COM STATUS DESATIVADO', size =(15, 3))],
    [sg.Button('STATUS DESATIVADO INICIO/FIM', size =(17, 2))],
    [sg.Button('INCIDENTE DE RDM', size =(16, 2))],
    [sg.Button('INDISPONIBILIDADE TOTAL', size =(15, 2))],
    [sg.Button('ABERTURA DE INCIDENTES', size =(15, 2))],
    [sg.Button('CONCLUIR INCIDENTES', size =(15, 2))],
    [sg.Button("Cancelar")]
]

#Centralizar
layout = [[sg.VPush()],
                [sg.Push(), sg.Column(layout,element_justification='c'), sg.Push()],
                [sg.VPush()]]

window = sg.Window('Desenvolver: Iran Santos', layout)

while True:
        event, values = window.read()

        if values and event == "Cancelar":
            break

        elif event == "STATUS DESATIVADO INICIO/FIM":
            sg.theme('DarkBlue4')     
            layout = [
                [sg.Text("PREENCHA OS CAMPOS ABAIXO:")],
                [sg.Text('OWNER', size =(15, 1)), sg.InputText()],
                [sg.Text('ISSUE', size =(15, 1)), sg.InputText()],
                [sg.Text('ADQUIRENTE', size =(15, 1)), sg.InputText()],
                [sg.Text('DATA/INICIO', size =(15, 1)), sg.InputText()],
                [sg.Submit('Enviar'), sg.Button("Cancelar")]
            ]

            layout = [[sg.VPush()],
                            [sg.Push(), sg.Column(layout,element_justification='c'), sg.Push()],
                            [sg.VPush()]]

            # VALUES 0 = OWNER
            # VALUES 1 = ISSUE
            # VALUES 2 = ADQUIRENTE
            # VALUES 3 = DATA/INICIO  
                        
            window = sg.Window('Desenvolver: Iran Santos', layout)

            while True:
                    event, values = window.read()
                    
                    if values and event == 'Enviar':  
                        Titulo = '[PAY] Adquirente com status desativado - ' + (values[2])
                        TINFRA = (values[1])
                        descricao = 'Incidente concluído. Informamos que o status do adquirente ' + (values[2]) +  ', no ambiente DTEF, encontra-se normalizado.'
                        descricao2 = 'Até o envio de encerramento deste incidente, o xCenter Pay não identificou mais nenhuma eventualidade no ambiente (monitoria).'
                        owner = (values[0])
                        impacto = 'Transações utilizando o adquirente ' + (values[2]) + ', não estão sendo completadas através dos servidores afetados, gerando timeouts financeiros e consequentemente vendas negadas.'
                        data = (values[3])
                        dataimpacto = data
                        acoes_corretivas = 'Inicialmente não houve ações de normalização pela equipe Cloud DTEF, sendo o evento uma possível oscilação na comunicação entre o adquirente ' + (values[2]) + ' e a LINX.'

                        webbrowser.open("https://grupolinx.sharepoint.com/sites/GestodeIncidentes-NOC/Lists/Disparo%20de%20Comunicado/Item/newifs.aspx?List=6fc06cb6%2Dd9a9%2D4255%2Db000%2Db0cf6f5e3d1e&Source=https%3A%2F%2Fgrupolinx%2Esharepoint%2Ecom%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%2520de%2520Comunicado%2FAllItems%2Easpx%3Fviewpath%3D%252Fsites%252FGestodeIncidentes%252DNOC%252FLists%252FDisparo%2520de%2520Comunicado%252FAllItems%252Easpx%26useFiltersInViewXml%3D1&RootFolder=%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%20de%20Comunicado&Web=879a6f25%2D038c%2D4e70%2Dac49%2D2ef1d7a7d509")
                        time.sleep(6)
                        pyautogui.write('TEF')
                        pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(3)
                        pyautogui.write('Concluído')
                        time.sleep(4)
                        pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(2)
                        pyautogui.write('PARCIAL/INTERMITENTE')
                        pyautogui.press(['tab']), pyautogui.press(['tab'])
                        pyautogui.write(Titulo)
                        pyautogui.press(['tab'])
                        pyautogui.write(TINFRA)
                        pyautogui.press(['tab'])
                        df=pd.DataFrame([descricao])
                        df.to_clipboard(index=False,header=False)
                        pyautogui.hotkey('ctrl', 'v')
                        df=pd.DataFrame([descricao2])
                        df.to_clipboard(index=False,header=False)
                        time.sleep(2)
                        pyautogui.hotkey('ctrl', 'v'), time.sleep(4)
                        pyautogui.press('backspace')
                        pyautogui.press('backspace')
                        time.sleep(2)
                        pyautogui.press(['tab']), time.sleep(3)
                        pyautogui.write(owner)
                        pyautogui.press(['tab'])
                        time.sleep(2)
                        pyautogui.write('N/A')
                        pyautogui.press(['tab'])
                        time.sleep(4)
                        df=pd.DataFrame([impacto])
                        df.to_clipboard(index=False,header=False)
                        pyautogui.hotkey('ctrl', 'v')
                        time.sleep(4)
                        pyautogui.press('backspace')
                        pyautogui.press('backspace')
                        time.sleep(2)
                        pyautogui.press(['tab'])
                        time.sleep(4)
                        df=pd.DataFrame([acoes_corretivas])
                        df.to_clipboard(index=False,header=False)
                        pyautogui.hotkey('ctrl', 'v')
                        pyautogui.press(['tab'])
                        time.sleep(4)
                        df=pd.DataFrame([values[3]])
                        df.to_clipboard(index=False,header=False)
                        pyautogui.hotkey('ctrl', 'v')
                        pyautogui.press(['tab'])
                        df=pd.DataFrame([values[3]])
                        df.to_clipboard(index=False,header=False)
                        pyautogui.hotkey('ctrl', 'v')
                    
                    elif event and values == "Cancelar":
                        break

                    elif event == window.close(): 
                        break

        elif event == 'INDISPONIBILIDADE TOTAL':
            sg.theme('DarkBlue4')
            layout = [
                    [sg.Text("PREENCHA OS CAMPOS ABAIXO:")],
                    [sg.Text('OWNER', size =(15, 1)), sg.InputText()],
                    [sg.Text('ISSUE', size =(15, 1)), sg.InputText()],
                    [sg.Text('SERVIDOR', size =(15, 1)), sg.InputText()],
                    [sg.Text('DATA/INICIO', size =(15, 1)), sg.InputText()],
                    [sg.Text('E-MAIL DO OWNER', size =(15, 1)), sg.InputText()],
                    [sg.Submit('Enviar'), sg.Button("Cancelar")]
                ]

            layout = [[sg.VPush()],
                                [sg.Push(), sg.Column(layout,element_justification='c'), sg.Push()],
                                [sg.VPush()]]

                # VALUES 0 = OWNER
                # VALUES 1 = ISSUE
                # VALUES 2 = ADQUIRENTE
                # VALUES 3 = DATA/INICIO  
                            
            window = sg.Window('Desenvolver: Iran Santos', layout)
            while True:
                event, values = window.read()
                        
                if values and event == 'Enviar':
                    Titulo = '[PAY] Indisponibilidade de transação aplicação DTEF-' + (values[2])
                    TINFRA = (values[1])
                    descricao = 'Informamos que identificamos através de nosso monitoramento xCenter Pay que a aplicação DTEF não está responsiva para o servidor ' + (values[2]) +  ', onde consequentemente inviabiliza as transações de todos os estabelecimentos cadastrados no servidor em questão.'
                    descricao2 = 'As evidências foram coletadas e encaminhadas para verificação da equipe Cloud TEF, assim que obtivermos informações de atualização, reportaremos no próximo comunicado.'
                    owner = (values[0])
                    impacto = 'Os clientes alocados no servidor ' + (values[2]) + ', ficam impossibilitados de realizar transações financeiras.'
                    data = (values[3])
                    dataimpacto = data
                    email_owner = (values[4])

                    webbrowser.open("https://grupolinx.sharepoint.com/sites/GestodeIncidentes-NOC/Lists/Disparo%20de%20Comunicado/Item/newifs.aspx?List=6fc06cb6%2Dd9a9%2D4255%2Db000%2Db0cf6f5e3d1e&Source=https%3A%2F%2Fgrupolinx%2Esharepoint%2Ecom%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%2520de%2520Comunicado%2FAllItems%2Easpx%3Fviewpath%3D%252Fsites%252FGestodeIncidentes%252DNOC%252FLists%252FDisparo%2520de%2520Comunicado%252FAllItems%252Easpx%26useFiltersInViewXml%3D1&RootFolder=%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%20de%20Comunicado&Web=879a6f25%2D038c%2D4e70%2Dac49%2D2ef1d7a7d509")
                    time.sleep(5)
                    pyautogui.write('TEF')
                    pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(3)
                    pyautogui.write('Primeiro Comunicado'), time.sleep(2)
                    pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(2)
                    pyautogui.write('TOTAL')
                    pyautogui.press(['tab']), pyautogui.press(['tab'])
                    df=pd.DataFrame([Titulo])
                    df.to_clipboard(index=False,header=False)
                    pyautogui.hotkey('ctrl', 'v')
                    pyautogui.press(['tab'])
                    pyautogui.write(TINFRA)
                    pyautogui.press(['tab'])
                    df=pd.DataFrame([descricao])
                    df.to_clipboard(index=False,header=False)
                    pyautogui.hotkey('ctrl', 'v')
                    df=pd.DataFrame([descricao2])
                    df.to_clipboard(index=False,header=False)
                    time.sleep(2)
                    pyautogui.hotkey('ctrl', 'v'), time.sleep(4)
                    pyautogui.press('backspace')
                    pyautogui.press('backspace')
                    time.sleep(2)
                    pyautogui.press(['tab']), time.sleep(3)
                    pyautogui.write(owner)
                    pyautogui.press(['tab'])
                    time.sleep(1)
                    pyautogui.write(email_owner)
                    pyautogui.press(['tab'])
                    time.sleep(2)
                    df=pd.DataFrame([impacto])
                    df.to_clipboard(index=False,header=False)
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(4)
                    pyautogui.press('backspace')
                    pyautogui.press('backspace')
                    time.sleep(2)
                    pyautogui.press(['tab'])
                    time.sleep(3)
                    df=pd.DataFrame([values[3]])
                    df.to_clipboard(index=False,header=False)
                    pyautogui.hotkey('ctrl', 'v')
                    pyautogui.press(['tab'])
                    df=pd.DataFrame([values[3]])
                    df.to_clipboard(index=False,header=False)
                    pyautogui.hotkey('ctrl', 'v')
                    

                elif event and values == "Cancelar":
                    break

                elif event == window.close(): 
                    break

        elif event == 'ADQUIRENTE COM STATUS DESATIVADO':
            sg.theme('DarkBlue4')     
            layout = [
                [sg.Text("PREENCHA OS CAMPOS ABAIXO:")],
                [sg.Text('OWNER', size =(15, 1)), sg.InputText()],
                [sg.Text('ISSUE', size =(15, 1)), sg.InputText()],
                [sg.Text('ADQUIRENTE', size =(15, 1)), sg.InputText()],
                [sg.Text('DATA/INICIO', size =(15, 1)), sg.InputText()],
                [sg.Submit('Enviar'), sg.Button("Cancelar")]
            ]

            layout = [[sg.VPush()],
                            [sg.Push(), sg.Column(layout,element_justification='c'), sg.Push()],
                            [sg.VPush()]]

            # VALUES 0 = OWNER
            # VALUES 1 = ISSUE
            # VALUES 2 = ADQUIRENTE
            # VALUES 3 = DATA/INICIO  
                        
            window = sg.Window('Desenvolver: Iran Santos', layout)

            while True:
                    event, values = window.read()
                    
                    if values and event == 'Enviar':  
                        Titulo = '[PAY] Adquirente com status desativado - ' + (values[2])
                        TINFRA = (values[1])
                        descricao = 'Identificamos que o adquirente ' + (values[2]) +  ' apresentou status desativado, nos ambiente DTEF, gerando elevação do percentual de timeouts financeiros e vendas negadas.'
                        descricao2 = 'As evidências foram coletadas e encaminhadas para verificação da equipe Cloud TEF, assim que obtivermos informações de atualização, reportaremos no próximo comunicado.'
                        owner = (values[0])
                        impacto = 'Transações utilizando o adquirente ' + (values[2]) + ' não estão sendo completadas através dos servidores afetados, gerando timeouts financeiros e consequentemente vendas negadas.'
                        data = (values[3])
                        dataimpacto = data

                        webbrowser.open("https://grupolinx.sharepoint.com/sites/GestodeIncidentes-NOC/Lists/Disparo%20de%20Comunicado/Item/newifs.aspx?List=6fc06cb6%2Dd9a9%2D4255%2Db000%2Db0cf6f5e3d1e&Source=https%3A%2F%2Fgrupolinx%2Esharepoint%2Ecom%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%2520de%2520Comunicado%2FAllItems%2Easpx%3Fviewpath%3D%252Fsites%252FGestodeIncidentes%252DNOC%252FLists%252FDisparo%2520de%2520Comunicado%252FAllItems%252Easpx%26useFiltersInViewXml%3D1&RootFolder=%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%20de%20Comunicado&Web=879a6f25%2D038c%2D4e70%2Dac49%2D2ef1d7a7d509")
                        time.sleep(6)
                        pyautogui.write('TEF')
                        pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(3)
                        pyautogui.write('Primeiro Comunicado')
                        time.sleep(4)
                        pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(2)
                        pyautogui.write('PARCIAL/INTERMITENTE')
                        pyautogui.press(['tab']), pyautogui.press(['tab'])
                        pyautogui.write(Titulo)
                        pyautogui.press(['tab'])
                        pyautogui.write(TINFRA)
                        pyautogui.press(['tab'])
                        df=pd.DataFrame([descricao])
                        df.to_clipboard(index=False,header=False)
                        pyautogui.hotkey('ctrl', 'v')
                        df=pd.DataFrame([descricao2])
                        df.to_clipboard(index=False,header=False)
                        time.sleep(2)
                        pyautogui.hotkey('ctrl', 'v'), time.sleep(4)
                        pyautogui.press('backspace')
                        pyautogui.press('backspace')
                        time.sleep(2)
                        pyautogui.press(['tab']), time.sleep(3)
                        pyautogui.write(owner)
                        pyautogui.press(['tab'])
                        time.sleep(2)
                        pyautogui.write('N/A')
                        pyautogui.press(['tab'])
                        time.sleep(4)
                        df=pd.DataFrame([impacto])
                        df.to_clipboard(index=False,header=False)
                        pyautogui.hotkey('ctrl', 'v')
                        time.sleep(4)
                        pyautogui.press('backspace')
                        pyautogui.press('backspace')
                        time.sleep(2)
                        pyautogui.press(['tab'])
                        time.sleep(3)
                        df=pd.DataFrame([values[3]])
                        df.to_clipboard(index=False,header=False)
                        pyautogui.hotkey('ctrl', 'v')
                        pyautogui.press(['tab'])
                        df=pd.DataFrame([values[3]])
                        df.to_clipboard(index=False,header=False)
                        pyautogui.hotkey('ctrl', 'v')
                        

                    elif event and values == "Cancelar":
                        break

                    elif event == window.close(): 
                        break
                       
        elif event == 'TIMEOUTS PARA ADQUIRENTES':
            #layout 2
            sg.theme('DarkBlue4')     
            layout = [
                [sg.Text("PREENCHA OS CAMPOS ABAIXO:")],
                [sg.Text('OWNER', size =(15, 1)), sg.InputText()],
                [sg.Text('ISSUE', size =(15, 1)), sg.InputText()],
                [sg.Text('ADQUIRENTE', size =(15, 1)), sg.InputText()],
                [sg.Text('DATA/INICIO', size =(15, 1)), sg.InputText()],
                [sg.Submit('Enviar'), sg.Button("Cancelar")]
            ]

            layout = [[sg.VPush()],
                            [sg.Push(), sg.Column(layout,element_justification='c'), sg.Push()],
                            [sg.VPush()]]

            # VALUES 0 = OWNER
            # VALUES 1 = ISSUE
            # VALUES 2 = ADQUIRENTE
            # VALUES 3 = DATA/INICIO  
                        
            window = sg.Window('Desenvolver: Iran Santos', layout)

            while True:
                    event, values = window.read()
                    
                    if values and event == 'Enviar': 

                            Titulo = '[PAY] Monitoramento DTEF - Timeouts Financeiros - ' + (values[2])
                            TINFRA = (values[1])
                            descricao = 'Informamos que identificamos a elevação no número de timeouts financeiros e vendas negadas no ambiente DTEF para o adquirente ' + (values[2]) +  ' as evidências foram coletadas e encaminhadas para a equipe Cloud TEF.'
                            descricao2 = 'Assim que obtivermos informações de atualização, reportaremos no próximo comunicado.'
                            owner = (values[0])
                            impacto = 'Parte das transações utilizando o adquirente ' + (values[2]) + ' não estão sendo concluídas em tempo hábil, gerando timeouts financeiros e consequentemente vendas negadas.'
                            data = (values[3])
                            dataimpacto = data

                            webbrowser.open("https://grupolinx.sharepoint.com/sites/GestodeIncidentes-NOC/Lists/Disparo%20de%20Comunicado/Item/newifs.aspx?List=6fc06cb6%2Dd9a9%2D4255%2Db000%2Db0cf6f5e3d1e&Source=https%3A%2F%2Fgrupolinx%2Esharepoint%2Ecom%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%2520de%2520Comunicado%2FAllItems%2Easpx%3Fviewpath%3D%252Fsites%252FGestodeIncidentes%252DNOC%252FLists%252FDisparo%2520de%2520Comunicado%252FAllItems%252Easpx%26useFiltersInViewXml%3D1&RootFolder=%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%20de%20Comunicado&Web=879a6f25%2D038c%2D4e70%2Dac49%2D2ef1d7a7d509")
                            time.sleep(6)
                            pyautogui.write('TEF')
                            pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(3)
                            pyautogui.write('Primeiro Comunicado')
                            time.sleep(4)
                            pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(2)
                            pyautogui.write('PARCIAL/INTERMITENTE')
                            pyautogui.press(['tab']), pyautogui.press(['tab'])
                            pyautogui.write(Titulo)
                            pyautogui.press(['tab'])
                            pyautogui.write(TINFRA)
                            pyautogui.press(['tab'])
                            df=pd.DataFrame([descricao])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            df=pd.DataFrame([descricao2])
                            df.to_clipboard(index=False,header=False)
                            time.sleep(2)
                            pyautogui.hotkey('ctrl', 'v'), time.sleep(4)
                            pyautogui.press('backspace')
                            pyautogui.press('backspace')
                            time.sleep(2)
                            pyautogui.press(['tab']), time.sleep(3)
                            pyautogui.write(owner)
                            pyautogui.press(['tab'])
                            time.sleep(2)
                            pyautogui.write('N/A')
                            pyautogui.press(['tab'])
                            time.sleep(4)
                            df=pd.DataFrame([impacto])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            time.sleep(4)
                            pyautogui.press('backspace')
                            pyautogui.press('backspace')
                            time.sleep(2)
                            pyautogui.press(['tab'])
                            time.sleep(3)
                            df=pd.DataFrame([values[3]])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            pyautogui.press(['tab'])
                            df=pd.DataFrame([values[3]])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')  
            
                    elif values and event == "Cancelar":
                        break

                    elif event and values == 'TIMEOUTS PARA ADQUIRENTES':
                        print('Em criação.')

                    elif event == window.close(): 
                        break

        elif event == 'INCIDENTE DE RDM':
            #layout 2
            sg.theme('DarkBlue4')     
            layout = [
                [sg.Text("PREENCHA OS CAMPOS ABAIXO:")],
                [sg.Text('OWNER', size =(15, 1)), sg.InputText()],
                [sg.Text('ISSUE', size =(15, 1)), sg.InputText()],
                [sg.Text('TITULO', size =(15, 1)), sg.InputText()],
                [sg.Text('DESCRICAO' , size =(15, 1)), sg.InputText()],
                [sg.Text(' Já vem com essa descricao: Recebemos relatos da equipe de P&D PayHub informando sobre o: DESCRICAO' , size =(55, 2)), sg.Text()],
                [sg.Text('O caso está sendo analisado pela equipe de P&D PayHub, assim que obtivermos maiores informações de atualização, reportaremos no próximo comunicado.' , size =(55, 2)), sg.Text()],
                [sg.Text('IMPACTO', size =(15, 1)), sg.InputText()],
                [sg.Text('DATA/INICIO', size =(15, 1)), sg.InputText()],
                [sg.Submit('Enviar'), sg.Button("Cancelar")]
            ]

            layout = [[sg.VPush()],
                            [sg.Push(), sg.Column(layout,element_justification='c'), sg.Push()],
                            [sg.VPush()]]

            # VALUES 0 = OWNER
            # VALUES 1 = ISSUE
            # VALUES 2 = TITULO
            # VALUES 3 = DESCRICAO
            # VALUES 4 = IMPACTO  
            # VALUES 5 = DATA/INICIO  
                        
            window = sg.Window('Desenvolver: Iran Santos', layout)

            while True:
                    event, values = window.read()
                    
                    if values and event == 'Enviar': 

                            Titulo = (values[2])
                            TINFRA = (values[1])
                            descricao = 'Recebemos relatos da equipe de P&D PayHub informando sobre o ' + (values[3])
                            descricao2 = 'O caso está sendo analisado pela equipe de P&D PayHub, assim que obtivermos maiores informações de atualização, reportaremos no próximo comunicado.'
                            owner = (values[0])
                            impacto = (values[4])
                            data = (values[5])
                            dataimpacto = data

                            webbrowser.open("https://grupolinx.sharepoint.com/sites/GestodeIncidentes-NOC/Lists/Disparo%20de%20Comunicado/Item/newifs.aspx?List=6fc06cb6%2Dd9a9%2D4255%2Db000%2Db0cf6f5e3d1e&Source=https%3A%2F%2Fgrupolinx%2Esharepoint%2Ecom%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%2520de%2520Comunicado%2FAllItems%2Easpx%3Fviewpath%3D%252Fsites%252FGestodeIncidentes%252DNOC%252FLists%252FDisparo%2520de%2520Comunicado%252FAllItems%252Easpx%26useFiltersInViewXml%3D1&RootFolder=%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%20de%20Comunicado&Web=879a6f25%2D038c%2D4e70%2Dac49%2D2ef1d7a7d509")
                            time.sleep(6)
                            pyautogui.write('TEF')
                            pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(3)
                            pyautogui.write('Primeiro Comunicado')
                            time.sleep(4)
                            pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(2)
                            pyautogui.write('PARCIAL/INTERMITENTE')
                            pyautogui.press(['tab']), pyautogui.press(['tab'])
                            pyautogui.write(Titulo)
                            pyautogui.press(['tab'])
                            pyautogui.write(TINFRA)
                            pyautogui.press(['tab'])
                            df=pd.DataFrame([descricao])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            df=pd.DataFrame([descricao2])
                            df.to_clipboard(index=False,header=False)
                            time.sleep(2)
                            pyautogui.hotkey('ctrl', 'v'), time.sleep(4)
                            pyautogui.press('backspace')
                            pyautogui.press('backspace')
                            time.sleep(2)
                            pyautogui.press(['tab']), time.sleep(3)
                            pyautogui.write(owner)
                            pyautogui.press(['tab'])
                            time.sleep(2)
                            pyautogui.write('N/A')
                            pyautogui.press(['tab'])
                            time.sleep(4)
                            df=pd.DataFrame([impacto])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            time.sleep(4)
                            pyautogui.press('backspace')
                            pyautogui.press('backspace')
                            time.sleep(2)
                            pyautogui.press(['tab'])
                            time.sleep(3)
                            df=pd.DataFrame([values[5]])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            pyautogui.press(['tab'])
                            df=pd.DataFrame([values[5]])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')

                    elif values and event == "Cancelar":
                        break

                    elif event == window.close(): 
                        break                         
         
        elif event == 'TIMEOUTS PARA ADQUIRENTES SITEF':
            #layout 2
            sg.theme('DarkBlue4')     
            layout = [
                [sg.Text("PREENCHA OS CAMPOS ABAIXO:")],
                [sg.Text('OWNER', size =(15, 1)), sg.InputText()],
                [sg.Text('ISSUE', size =(15, 1)), sg.InputText()],
                [sg.Text('ADQUIRENTE', size =(15, 1)), sg.InputText()],
                [sg.Text('DATA/INICIO', size =(15, 1)), sg.InputText()],
                [sg.Submit('Enviar'), sg.Button("Cancelar")]
            ]

            layout = [[sg.VPush()],
                            [sg.Push(), sg.Column(layout,element_justification='c'), sg.Push()],
                            [sg.VPush()]]

            # VALUES 0 = OWNER
            # VALUES 1 = ISSUE
            # VALUES 2 = ADQUIRENTE
            # VALUES 3 = DATA/INICIO  
                        
            window = sg.Window('Desenvolver: Iran Santos', layout)

            while True:
                    event, values = window.read()
                    
                    if values and event == 'Enviar': 

                            Titulo = '[PAY] Monitoramento SITEF - Timeouts Financeiros - ' + (values[2])
                            TINFRA = (values[1])
                            descricao = 'Informamos que identificamos a elevação no percentual de timeouts financeiros e vendas negadas no ambiente SITEF para o adquirente ' + (values[2]) +  ' as evidências foram coletadas e encaminhadas para a equipe Cloud SITEF.'
                            descricao2 = 'Assim que obtivermos informações de atualização, reportaremos no próximo comunicado.'
                            owner = (values[0])
                            impacto = 'Parte das transações utilizando o adquirente ' + (values[2]) + ' não estão sendo concluídas em tempo hábil, gerando timeouts financeiros e consequentemente vendas negadas.'
                            data = (values[3])
                            dataimpacto = data

                            webbrowser.open("https://grupolinx.sharepoint.com/sites/GestodeIncidentes-NOC/Lists/Disparo%20de%20Comunicado/Item/newifs.aspx?List=6fc06cb6%2Dd9a9%2D4255%2Db000%2Db0cf6f5e3d1e&Source=https%3A%2F%2Fgrupolinx%2Esharepoint%2Ecom%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%2520de%2520Comunicado%2FAllItems%2Easpx%3Fviewpath%3D%252Fsites%252FGestodeIncidentes%252DNOC%252FLists%252FDisparo%2520de%2520Comunicado%252FAllItems%252Easpx%26useFiltersInViewXml%3D1&RootFolder=%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%20de%20Comunicado&Web=879a6f25%2D038c%2D4e70%2Dac49%2D2ef1d7a7d509")
                            time.sleep(6)
                            pyautogui.write('TEF')
                            pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(3)
                            pyautogui.write('Primeiro Comunicado')
                            time.sleep(4)
                            pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(2)
                            pyautogui.write('PARCIAL/INTERMITENTE')
                            pyautogui.press(['tab']), pyautogui.press(['tab'])
                            pyautogui.write(Titulo)
                            pyautogui.press(['tab'])
                            pyautogui.write(TINFRA)
                            pyautogui.press(['tab'])
                            df=pd.DataFrame([descricao])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            df=pd.DataFrame([descricao2])
                            df.to_clipboard(index=False,header=False)
                            time.sleep(2)
                            pyautogui.hotkey('ctrl', 'v'), time.sleep(4)
                            pyautogui.press('backspace')
                            pyautogui.press('backspace')
                            time.sleep(2)
                            pyautogui.press(['tab']), time.sleep(3)
                            pyautogui.write(owner)
                            pyautogui.press(['tab'])
                            time.sleep(2)
                            pyautogui.write('N/A')
                            pyautogui.press(['tab'])
                            time.sleep(4)
                            df=pd.DataFrame([impacto])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            time.sleep(4)
                            pyautogui.press('backspace')
                            pyautogui.press('backspace')
                            time.sleep(2)
                            pyautogui.press(['tab'])
                            time.sleep(3)
                            df=pd.DataFrame([values[3]])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            pyautogui.press(['tab'])
                            df=pd.DataFrame([values[3]])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v') 
            
                    elif values and event == "Cancelar":
                        break

                    elif event == window.close(): 
                        break

        elif event == 'ABERTURA DE INCIDENTES':
            #layout 2
            sg.theme('DarkBlue4')     
            layout = [
                [sg.Text("PREENCHA OS CAMPOS ABAIXO:")],
                [sg.Text('OWNER', size =(15, 1)), sg.InputText()],
                [sg.Text('ISSUE', size =(15, 1)), sg.InputText()],
                [sg.Text('TITULO', size =(15, 1)), sg.InputText()],
                [sg.Text('DESCRICAO', size =(15, 1)), sg.InputText()],
                [sg.Text('Assim que obtivermos informações de atualização, reportaremos no próximo comunicado.' , size =(55, 2)), sg.Text()],
                [sg.Text('IMPACTO', size =(15, 1)), sg.InputText()],
                [sg.Text('DATA/INICIO', size =(15, 1)), sg.InputText()],
                [sg.Submit('Enviar'), sg.Button("Cancelar")]
            ]

            layout = [[sg.VPush()],
                            [sg.Push(), sg.Column(layout,element_justification='c'), sg.Push()],
                            [sg.VPush()]]

            # VALUES 0 = OWNER
            # VALUES 1 = ISSUE
            # VALUES 2 = TITULO
            # VALUES 3 = DESCRICAO 1
            # VALUES 4 = IMPACTO
            # VALUES 5 = DATA/INICIO  
                        
            window = sg.Window('Desenvolver: Iran Santos', layout)

            while True:
                    event, values = window.read()
                    
                    if values and event == 'Enviar': 

                            Titulo = (values[2])
                            TINFRA = (values[1])
                            descricao = (values[3])
                            descricao2 = 'Assim que obtivermos informações de atualização, reportaremos no próximo comunicado.'
                            owner = (values[0])
                            impacto = (values[4])
                            data = (values[5])
                            dataimpacto = data

                            webbrowser.open("https://grupolinx.sharepoint.com/sites/GestodeIncidentes-NOC/Lists/Disparo%20de%20Comunicado/Item/newifs.aspx?List=6fc06cb6%2Dd9a9%2D4255%2Db000%2Db0cf6f5e3d1e&Source=https%3A%2F%2Fgrupolinx%2Esharepoint%2Ecom%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%2520de%2520Comunicado%2FAllItems%2Easpx%3Fviewpath%3D%252Fsites%252FGestodeIncidentes%252DNOC%252FLists%252FDisparo%2520de%2520Comunicado%252FAllItems%252Easpx%26useFiltersInViewXml%3D1&RootFolder=%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%20de%20Comunicado&Web=879a6f25%2D038c%2D4e70%2Dac49%2D2ef1d7a7d509")
                            time.sleep(6)
                            pyautogui.write('TEF')
                            pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(3)
                            pyautogui.write('Primeiro Comunicado')
                            time.sleep(4)
                            pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(2)
                            pyautogui.write('PARCIAL/INTERMITENTE')
                            pyautogui.press(['tab']), pyautogui.press(['tab'])
                            df=pd.DataFrame([Titulo])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            pyautogui.press(['tab'])
                            pyautogui.write(TINFRA)
                            pyautogui.press(['tab'])
                            df=pd.DataFrame([descricao])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            df=pd.DataFrame([descricao2])
                            df.to_clipboard(index=False,header=False)
                            time.sleep(2)
                            pyautogui.hotkey('ctrl', 'v'), time.sleep(4)
                            pyautogui.press('backspace')
                            pyautogui.press('backspace')
                            time.sleep(2)
                            pyautogui.press(['tab']), time.sleep(3)
                            pyautogui.write(owner)
                            pyautogui.press(['tab'])
                            time.sleep(2)
                            pyautogui.write('N/A')
                            pyautogui.press(['tab'])
                            time.sleep(4)
                            df=pd.DataFrame([impacto])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            time.sleep(4)
                            pyautogui.press('backspace')
                            pyautogui.press('backspace')
                            time.sleep(2)
                            pyautogui.press(['tab'])
                            time.sleep(3)
                            df=pd.DataFrame([values[5]])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            pyautogui.press(['tab'])
                            df=pd.DataFrame([values[5]])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v') 
            
                    elif values and event == "Cancelar":
                        break

                    elif event == window.close(): 
                        break

        elif event == 'CONCLUIR INCIDENTES': 
            sg.theme('DarkBlue4')
            layout = [
                    [sg.Text("PREENCHA OS CAMPOS ABAIXO:")],
                    [sg.Text('ISSUE', size =(15, 1)), sg.InputText()],
                    [sg.Text('ADQUIRENTE', size =(15, 1)), sg.InputText()],
                    [sg.Text('DATA/FIM', size =(15, 1)), sg.InputText()],
                    [sg.Text('CONCLUSAO', size =(15, 1)), sg.InputText()],
                    [sg.Text('ACOES CORRETIVAS', size =(15, 2)), sg.InputText()],
                    [sg.Submit('Enviar'), sg.Button("Cancelar")]
                ]

            layout = [[sg.VPush()],
                                                    [sg.Push(), sg.Column(layout,element_justification='c'), sg.Push()],
                                                    [sg.VPush()]]

                
                    # VALUES 0 = ISSUE
                    # VALUES 1 = ADQUIRENTE
                    # VALUES 2 = DATA/FIM
                    # VALUES 3 = TEXTO CONCLUSAO
                    # VALUES 4 = ACOES CORRETIVAS

            window = sg.Window('Desenvolver: Iran Santos', layout)
            while True:
                event, values = window.read()
                                            
                if values and event == 'Enviar':

                            descricao = (values[3])
                            descricao2 = 'Ressaltamos que o evento foi momentâneo conforme o período evidenciado, ocasionando a elevação de tempos médios de transações e logo houve a normalização do ambiente.'
                            data_termino = (values[2])
                            issue = (values[0])
                            conclusao = (values[4])

                            webbrowser.open("https://grupolinx.sharepoint.com/sites/GestodeIncidentes-NOC/Lists/Disparo%20de%20Comunicado/AllItems.aspx?viewpath=%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%20de%20Comunicado%2FAllItems%2Easpx&useFiltersInViewXml=1")
                            time.sleep(6)
                            pyautogui.press(['tab']), pyautogui.press(['tab']), time.sleep(1)
                            pyautogui.press(['tab']), pyautogui.press(['tab']), time.sleep(1)
                            pyautogui.press(['tab']), pyautogui.press(['tab']), time.sleep(1)
                            pyautogui.press(['tab']), time.sleep(1)
                            pyautogui.press(['tab']), time.sleep(1)
                            pyautogui.press(['tab']), time.sleep(1), pyautogui.press(['tab']), time.sleep(1)
                            time.sleep(2)
                            pyautogui.write(issue) 
                            pyautogui.press(['enter']) 
                            time.sleep(2)
                            pyautogui.press(['enter']) 
                            time.sleep(6)
                            pyautogui.press(['tab']), time.sleep(3), pyautogui.press(['tab']), time.sleep(3)
                            pyautogui.press(['enter']), time.sleep(3) #EDITAR ITEM
                            pyautogui.press(['tab']), time.sleep(3), pyautogui.press(['tab']), time.sleep(3)
                            pyautogui.write('Con'), time.sleep(3)
                            pyautogui.press(['tab']), time.sleep(1),
                            pyautogui.press(['tab']), time.sleep(1), pyautogui.press(['tab']), time.sleep(1), pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(1), pyautogui.press(['tab'])
                            df=pd.DataFrame([descricao]), time.sleep(4)
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v'), time.sleep(2)
                            df=pd.DataFrame([descricao2]), time.sleep(4)
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v'), time.sleep(2)
                            pyautogui.press(['tab']), time.sleep(3)
                            pyautogui.press(['tab'])
                            time.sleep(2)
                            pyautogui.press(['tab'])
                            pyautogui.write(conclusao)
                            time.sleep(2)
                            pyautogui.press(['tab'])
                            time.sleep(3)
                            pyautogui.press(['tab'])
                            time.sleep(2)
                            pyautogui.press(['tab'])
                            df=pd.DataFrame([data_termino])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v'), time.sleep(2)###

                elif event and values == "Cancelar":
                    break

                elif event == window.close(): 
                    break    

        elif event == 'TIMEOUTS INICIO/FIM': 
            
            #layout 2
            sg.theme('DarkBlue4')     
            layout = [
                [sg.Text("PREENCHA OS CAMPOS ABAIXO:")],
                [sg.Text('OWNER', size =(15, 1)), sg.InputText()],
                [sg.Text('ISSUE', size =(15, 1)), sg.InputText()],
                [sg.Text('ADQUIRENTE', size =(15, 1)), sg.InputText()],
                [sg.Text('DATA/INICIO', size =(15, 1)), sg.InputText()],
                [sg.Submit('Enviar'), sg.Button("Cancelar")]
            ]

            layout = [[sg.VPush()],
                            [sg.Push(), sg.Column(layout,element_justification='c'), sg.Push()],
                            [sg.VPush()]]

            # VALUES 0 = OWNER
            # VALUES 1 = ISSUE
            # VALUES 2 = ADQUIRENTE
            # VALUES 3 = DATA/INICIO  


            window = sg.Window('Desenvolver: Iran Santos', layout)

            while True:
                    event, values = window.read()
                    
                    if values and event == 'Enviar': 

                            Titulo = '[PAY] Monitoramento DTEF - Timeouts Financeiros - ' + (values[2])
                            TINFRA = (values[1])
                            descricao = 'Incidente concluído. Informamos que o número de timeouts financeiros no ambiente DTEF se encontra normalizado (sem reincidências) para as transações do adquirente ' + (values[2]) + '.'
                            descricao2 = 'Até o envio de encerramento deste incidente, o xCenter Pay não identificou mais nenhuma eventualidade no ambiente (monitoria).'
                            owner = (values[0])
                            impacto = 'Parte das transações utilizando o adquirente ' + (values[2]) + ' não estão sendo concluídas em tempo hábil, gerando timeouts financeiros e consequentemente vendas negadas.'
                            data = (values[3])
                            dataimpacto = data
                            acoes_corretivas = 'Inicialmente não houve ações de normalização pela equipe Cloud DTEF, sendo o evento uma possível oscilação na comunicação entre o adquirente ' + (values[2]) + ' e a LINX.'

                            webbrowser.open("https://grupolinx.sharepoint.com/sites/GestodeIncidentes-NOC/Lists/Disparo%20de%20Comunicado/Item/newifs.aspx?List=6fc06cb6%2Dd9a9%2D4255%2Db000%2Db0cf6f5e3d1e&Source=https%3A%2F%2Fgrupolinx%2Esharepoint%2Ecom%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%2520de%2520Comunicado%2FAllItems%2Easpx%3Fviewpath%3D%252Fsites%252FGestodeIncidentes%252DNOC%252FLists%252FDisparo%2520de%2520Comunicado%252FAllItems%252Easpx%26useFiltersInViewXml%3D1&RootFolder=%2Fsites%2FGestodeIncidentes%2DNOC%2FLists%2FDisparo%20de%20Comunicado&Web=879a6f25%2D038c%2D4e70%2Dac49%2D2ef1d7a7d509")
                            time.sleep(6)
                            pyautogui.write('TEF')
                            pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(3)
                            pyautogui.write('Concluído')
                            time.sleep(4)
                            pyautogui.press(['tab']), time.sleep(2), pyautogui.press(['tab']), time.sleep(2)
                            pyautogui.write('PARCIAL/INTERMITENTE')
                            pyautogui.press(['tab']), pyautogui.press(['tab'])
                            pyautogui.write(Titulo)
                            pyautogui.press(['tab'])
                            pyautogui.write(TINFRA)
                            pyautogui.press(['tab'])
                            df=pd.DataFrame([descricao])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            df=pd.DataFrame([descricao2])
                            df.to_clipboard(index=False,header=False)
                            time.sleep(2)
                            pyautogui.hotkey('ctrl', 'v'), time.sleep(4)
                            pyautogui.press('backspace')
                            pyautogui.press('backspace')
                            time.sleep(2)
                            pyautogui.press(['tab']), time.sleep(3)
                            pyautogui.write(owner)
                            pyautogui.press(['tab'])
                            time.sleep(2)
                            pyautogui.write('N/A')
                            pyautogui.press(['tab'])
                            time.sleep(4)
                            df=pd.DataFrame([impacto])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            time.sleep(4)
                            pyautogui.press('backspace')
                            pyautogui.press('backspace')
                            time.sleep(2)
                            pyautogui.press(['tab'])
                            time.sleep(4)
                            df=pd.DataFrame([acoes_corretivas])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v'), time.sleep(3)
                            pyautogui.press('backspace')
                            pyautogui.press('backspace')
                            time.sleep(2)
                            pyautogui.press(['tab'])
                            time.sleep(4)
                            df=pd.DataFrame([values[3]])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')
                            pyautogui.press(['tab'])
                            df=pd.DataFrame([values[3]])
                            df.to_clipboard(index=False,header=False)
                            pyautogui.hotkey('ctrl', 'v')

                    elif values and event == "Cancelar":
                        break

                    elif event == window.close(): 
                        break 

        elif event == window.close(): 
            break

window.close()    