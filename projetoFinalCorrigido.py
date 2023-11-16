import pandas as pd
from datetime import datetime as dt

df_acesso = pd.read_excel('c:/Users/Isadora/Desktop/pythonfinal/Acessos.xlsx')
df_chaves = pd.read_excel('c:/Users/Isadora/Desktop/pythonfinal/disponibilidade.xlsx')
df_movimento = pd.read_excel('c:/Users/Isadora/Desktop/pythonfinal/movimentos.xlsx')

userId = int(input("Entre com o seu ID: "))
# ID 123 é o único que acessa as funções de administrador, portanto é o unico que consegue adicionar usuário ou editar permissões.

def administrador(df_acesso, userId):
    adm = input("Deseja adicionar novo usuário ou editar as permissões de chaves? (adicionar/editar): ")
    if adm == 'adicionar':
        while True:
            novoID = input("Entre com o novo ID da pessoa que deseja adicionar: ")
            newName = input("Entre com o nome da pessoa que deseja adicionar: ")
            linha_autorizado = pd.DataFrame({'ID': [novoID], 'Nome': [newName], 'Chave A': ["S"], 'Chave B': ["S"], 'Chave C': ["S"], 'Chave D': ["S"], 'Chave E': ["S"], 'Adm': ["N"]})
            df_acesso = pd.concat([df_acesso, linha_autorizado], ignore_index=True) #gravar nova linha no excel "movimentos"
            df_acesso.to_excel('c:/Users/Isadora/Desktop/pythonfinal/Acessos.xlsx', index=False)
            adicionar = input("Deseja adicionar outra pessoa? (sim/não): ")
            if adicionar == 'sim':
                continue
            else:
                break
    elif adm == 'editar':
        print(df_acesso['ID'].values) #faz o print dos ID existentes para que o administrador possa escolher
        while True:
            editar_user = int(input("Entre com o ID do usuário que deseja editar: "))
            if editar_user in df_acesso['ID'].values:
                print("Permissões de chaves atuais para o usuário:")
                print(df_acesso[df_acesso['ID'] == editar_user]) #faz o print das atuais permissões do usuario escolhido pelo administrador
                nova_permissao = input("Qual chave deseja editar? (Chave A/Chave B/Chave C/Chave D/Chave E): ")
                if nova_permissao in ['Chave A', 'Chave B', 'Chave C', 'Chave D', 'Chave E']:
                    indx = df_acesso[(df_acesso['ID'] == editar_user) & (df_acesso[nova_permissao] == 'S')].index
                    if len(indx) > 0:
                        editado = df_acesso.at[indx[0], nova_permissao]
                        if editado == "S": #Tira o acesso da chave escolhida
                            df_acesso.at[indx[0], nova_permissao] = 'N'
                        else:
                            df_acesso.at[indx[0], nova_permissao] = 'S' #Dá acesso para a chave escolhida
                        df_acesso.to_excel('c:/Users/Isadora/Desktop/pythonfinal/Acessos.xlsx', index=False)
                        print("Permissões de chaves atualizadas com sucesso.")
                        editar_novamente = input("Deseja fazer outra edição? (sim/não): ")
                        if editar_novamente == 'sim':
                            continue
                        else:
                            break
                else:
                    print("Chave inválida. Escolha entre 'Chave A', 'Chave B', 'Chave C', 'Chave D' ou 'Chave E'.")
                    continue
            else:
                print("ID do usuário não encontrado.")
                continue
    else:
        print("Opção inválida. Escolha 'adicionar' ou 'editar'.")
    encerrar_programa = input("Deseja encerrar o programa? (sim/não): ") #Dá a opção de encerrar o programa
    if encerrar_programa == 'não':
        verificarAcesso(df_acesso, df_chaves, df_movimento, userId) #volta o programa para o início
    else:
        print("Obrigada e até logo!") #Encerra o programa

def acao(df_acesso, df_chaves, df_movimento, userId):
    nome = df_acesso[df_acesso['ID'] == userId]['Nome'].values[0] #Busca o nome do usuário a partir do ID inserido para posteriormente gravar o movimento
    data_hora = dt.now().strftime('%Y-%m-%d %H:%M:%S') #Graava no excel a data e a hora do movimento
    while True:
        levRet = input("Deseja levantar ou retornar uma chave? (levantar/retornar): ")
        if levRet == 'levantar':
            while True:
                chave = input("Entre com o nome da chave (A, B, C, D, E): ")
                if chave in df_chaves['Chave'].values:
                    idx = df_chaves[df_chaves['Chave'] == chave].index[0]
                    disponibilidade = df_chaves.at[idx, 'Livre'] #verifica se a chave está disponível
                    if disponibilidade == "S":
                        print("Chave disponível, já pode ser levantada.")
                        df_chaves.at[idx, 'Livre'] = 'N'
                        df_chaves.to_excel('c:/Users/Isadora/Desktop/pythonfinal/disponibilidade.xlsx', index=False)
                        linha_levantar = pd.DataFrame({'ID': [userId], 'Chave': [chave], 'Nome': [nome], 'Ação': ["L"], 'Data': [data_hora]})
                        df_movimento = pd.concat([df_movimento, linha_levantar], ignore_index=True) #Grava no excel movimento que a chava foi levantada, a hora e por quem
                        df_movimento.to_excel('c:/Users/Isadora/Desktop/pythonfinal/movimentos.xlsx', index=False)
                    else:
                        print("Chave em uso.")
                else:
                    print("Chave inexistente")
                outraChave = input("Deseja pegar outra chave? (sim/não) ") 
                if outraChave == 'sim':
                    continue
                else:
                    break
        elif levRet == 'retornar':
            print(df_movimento[df_movimento['ID'] == userId]) #Faz o print das chaves que o usuário possui consigo
            devolver = input("Entre com a identificação da chave que deseja retornar (A, B, C, D, E): ")
            if devolver in df_chaves['Chave'].values:
                indx = df_chaves[df_chaves['Chave'] == devolver].index[0]
                indisponivel = df_chaves.at[indx, 'Livre']
                if indisponivel == "N":
                    print("Chave retornada.")
                    df_chaves.at[indx, 'Livre'] = 'S' #Ao retornar uma chave, o programa automaticamente deixa ela disponivel no excel "disponibilidade"
                    df_chaves.to_excel('c:/Users/Isadora/Desktop/pythonfinal/disponibilidade.xlsx', index=False)
                    linha_retornar = pd.DataFrame({'ID': [userId], 'Chave': [devolver], 'Nome': [nome], 'Ação': ["R"], 'Data': [data_hora]})
                    df_movimento = pd.concat([df_movimento, linha_retornar], ignore_index=True)
                    df_movimento.to_excel('c:/Users/Isadora/Desktop/pythonfinal/movimentos.xlsx', index=False)
                else:
                    df_chaves.at[indx, 'Livre'] = 'S' #Corrigi bug, caso no excel ela esteja como disponivel qdo na verdade está com alguém
                    df_chaves.to_excel('c:/Users/Isadora/Desktop/pythonfinal/disponibilidade.xlsx', index=False)
                    linha_retornar = pd.DataFrame({'ID': [userId], 'Chave': [devolver], 'Nome': [nome], 'Ação': ["R"], 'Data': [data_hora]})
                    df_movimento = pd.concat([df_movimento, linha_retornar], ignore_index=True)
                    df_movimento.to_excel('c:/Users/Isadora/Desktop/pythonfinal/movimentos.xlsx', index=False)
            else:
                print("Chave inexistente")
                retornarOutra = input("Deseja retornar outra chave? (sim/não) ") 
                if retornarOutra == 'sim':
                    continue
                else:
                    break
        else:
            print("Ação inválida. Escolha 'levantar' ou 'retornar'.")
        the_end = input("Deseja encerrar o programa? (sim/não): ") 
        if the_end == 'não':
            verificarAcesso(df_acesso, df_chaves, df_movimento, userId) #Volta o programa do início. 
        else:
            print("Obrigada e até logo!")
        break

def verificarAcesso(df_acesso, df_chaves, df_movimento, userId):
    if userId in df_acesso['ID'].values:
        print("Acesso permitido.")
        if df_acesso[df_acesso['ID'] == userId]['Adm'].values[0] == 'S': #vefirica se o usuário possui permissão de administrador
            perguntar = input("Deseja acessar funções administrativas? (sim/não): ")
            if perguntar == 'sim':
                administrador(df_acesso, userId) #Acessa as funções de administrador
            else:
                acao(df_acesso, df_chaves, df_movimento, userId) #Acessa as funções de levantar ou retornar uma chave
        else:
            acao(df_acesso, df_chaves, df_movimento, userId) #Se o usuário não for administrador, vai direto para a função levantar ou retornar uma chave
    else:
        print("Acesso negado. ID não encontrado ou sem permissão.")

verificarAcesso(df_acesso, df_chaves, df_movimento, userId)
#Atividade Isadora Silva