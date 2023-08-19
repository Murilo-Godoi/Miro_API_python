import pandas as pd
import requests
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import simpledialog
import datetime
import numpy as np
import ctypes


# ----- funções da interface
def get_board_id_from_user_input(user_input):  # pega o id da lousa a partir do link
    splited_input = str(user_input).split("/")
    splited_input = [
        i for i in splited_input if "=" in i
    ]  # o id da lousa sempre tem um '=' no final

    if len(splited_input) != 1:
        ctypes.windll.user32.MessageBoxW(
            0, "Essa entrada nao corresponde a uma lousa existente!", "Miro API", 1
        )
        return 0

    # requisição qualquer só pra testar conexão
    board_id = splited_input[0]
    url = f"https://api.miro.com/v2/boards/{board_id}/members"
    headers = {"accept": "application/json", "authorization": f"Bearer {access_token}"}
    print(url)
    response = requests.get(url, headers=headers)
    print(response.text)

    if response.status_code == 401:
        ctypes.windll.user32.MessageBoxW(
            0,
            "Token de acesso invalido (utilize o botão de ajuda se necessário)",
            "Miro API",
            1,
        )
        return 0

    elif response.status_code == 200:
        ctypes.windll.user32.MessageBoxW(0, "Lousa selecionada!", "Miro API", 1)
        return board_id.strip()

    ctypes.windll.user32.MessageBoxW(
        0,
        "Nao foi possivel conectar-se (utilize as instruções se necessário)",
        "Miro API",
        1,
    )
    return 0


def retrieve_input(textBox):  # função para pegar valor da textBox
    global board_id
    userInput = textBox.get("1.0", "end-1c")
    board_id = get_board_id_from_user_input(userInput)


def read_datase(excel_path, from_credentials=False):
    try:
        database = pd.read_excel(excel_path, names=["dados", "classificacoes"])
        database = database[database["dados"].notna()]

        if from_credentials:
            ctypes.windll.user32.MessageBoxW(
                0, "Usando a base de dados da ultima seção!", "Miro API", 1
            )
        else:
            df_db_path = pd.DataFrame(
                [{"Database": excel_path}]
            )  # que jeito feio de salvar o token
            df_db_path.to_csv("last_database.csv", index=False)

        database_not_null = database.fillna("")
        ctypes.windll.user32.MessageBoxW(0, "Base de dados selecionada!", "Miro API", 1)
        return (
            dict(zip(database_not_null["dados"], database_not_null["classificacoes"])),
            database,
        )

    except PermissionError:
        ctypes.windll.user32.MessageBoxW(
            0,
            "Algo deu errado ao tentar ler a base de dados, verifique se o arquivo está aberto",
            "Miro API",
            1,
        )

    except ValueError:
        ctypes.windll.user32.MessageBoxW(
            0,
            "Algo deu errado ao tentar ler a base de dados, verifique se o arquivo fornecido é um excel",
            "Miro API",
            1,
        )

    except IndexError:
        ctypes.windll.user32.MessageBoxW(
            0,
            "Algo deu errado ao tentar ler a base de dados, verificque se o arquivo fornecido tem pelo menos duas colunas",
            "Miro API",
            1,
        )

    except Exception as e:
        print(e)
        ctypes.windll.user32.MessageBoxW(
            0, "Algo deu errado ao tentar ler a base de dados", "Miro API", 1
        )

    return 0


# tutoriais e inputs fora da interface
def show_select_database_dialog():
    global card_tags, database

    root = Tk()
    excel_path = filedialog.askopenfilename(
        title="Selecione a base de dados"
    )  # no momento, essa base de dados deve ser um arquivo excel onde a primeira coluna representam os dados e a segunda suas classificacoes
    card_tags, database = read_datase(excel_path)

    root.destroy()


def show_access_token_input():
    global access_token

    access_token = simpledialog.askstring(
        title="Miro API",
        prompt="Forneça o token de acesso:\n(use o botao de ajuda se necessário)",
    )

    if not access_token:
        ctypes.windll.user32.MessageBoxW(0, "Token inválido!", "Miro API", 1)
    else:
        df_token = pd.DataFrame(
            [{"Access_Token": access_token}]
        )  # que jeito feio de salvar o token
        df_token.to_csv("credentials.csv", index=False)


def show_instrucoes_de_uso():
    str_tutorial = """
INSTRUÇÕES

1. Para começar, insira o link da lousa na caixa de texto e clique em conectar
(a lousa deve estar dentro do time da ultima seção, caso contrário é necessário fornecer o token do time novamente)

2. Por padrão, será utilizado o arquivo "Base de dados.xlsx", salvo na mesma pasta do programa, como base de dados.

3. Caso queira alterar, use a opção "Selecionar base de dados".\n(a base de dados deve ser um arquivo excel, onde a primeira coluna será utilizada como anotação dentro do stick note e segunda será sua respectiva tag)

4. Para enviar dados da base para o Miro, use a opção "Importar Cards"

5. Para puxar os dados e classificacoes do Miro para a base, use a opção "Exportar Cards".

6. Os dados extraídos do Miro são salvos na mesma pasta do programa, em um arquivo chamado "Nova Classificação.xlsx".

7. Ao exportar, também será gerado um excel mostrando quais classificacoes foram alteradas

Caso ainda nao tenha habilitado a lousa para o time em que ela está inserida, use o botão de "ajuda" na interface principal do programa
"""
    ctypes.windll.user32.MessageBoxW(0, str_tutorial, "Miro API", 1)


def show_enable_board_tutorial():
    str_tutorial = """
HABILITAR API PARA UM TIME DO MIRO

1. Dentro do Miro, clique em sua foto, no canto superior direito.

2. Em seguida, clique em "Meus aplicativos" e depois em "Criar novo aplicativo". De um nome, escolha um time e clique em "criar aplicativo". 

3. Desça até "Permissões", marque as opções "boards:read" e "boards:write" e clique em "instalar aplicativo e pegar Token" 

4. Escolha o mesmo time e cique em "Adicionar". Em seguida será fornecido o token de acesso.

5. Clique no botão "Alterar token" da interface desse programa e cole o token de acesso.

6. Agora a API está habilitada para esse time e basta fornecer o token e o link de qualquer lousa dentro desse time para conectar-se 

7. Por padrão, ao iniciar o programa será utilizado o token da lousa usada na última seção.

"""
    ctypes.windll.user32.MessageBoxW(0, str_tutorial, "Miro API", 1)


# ----- funções para pegar cards e tags (e seus respectivos IDs) que acabaram de subir e os que ja estão na lousa
def get_tags_on_board():
    url = f"https://api.miro.com/v2/boards/{board_id}/tags?limit=50"

    headers = {"accept": "application/json", "authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)

    response_json = response.json()
    tag_ids = {}
    for result in response_json["data"]:
        name = result["title"]
        id = result["id"]

        tag_ids[name] = id

    print(tag_ids)
    return tag_ids


def get_cards_on_board():
    url = f"https://api.miro.com/v2/boards/{board_id}/items?limit=50&type=sticky_note"

    headers = {"accept": "application/json", "authorization": f"Bearer {access_token}"}

    response = requests.get(url, headers=headers)
    response_json = response.json()

    card_ids = {}
    for result in response_json["data"]:
        name = result["data"]["content"]
        id = result["id"]

        card_ids[id] = name

    return card_ids


# ----- enviando dados da base para o Miro
def post_card_and_tags():
    if not card_tags:  # significa que nao leu uma base de dados
        ctypes.windll.user32.MessageBoxW(
            0, "Por favor selecione uma base de dados", "Miro API", 1
        )
        return 0

    if not board_id:  # significa que nao leu uma base de dados
        ctypes.windll.user32.MessageBoxW(
            0, "Por favor conecte-se a uma lousa", "Miro API", 1
        )
        return 0

    create_cards()
    create_tags()

    card_ids = get_cards_on_board()
    tag_ids = get_tags_on_board()

    attach_tag_to_card(card_ids, tag_ids)


def create_cards():
    card_ids = {}

    # posição dos stick notes
    x = 0
    y = 0
    for card in card_tags:
        url = f"https://api.miro.com/v2/boards/{board_id}/sticky_notes"

        payload = {
            "data": {"content": f"{card}"},
            "position": {"origin": "center", "x": x, "y": y},
        }
        headers = {
            "accept": "application/json",
            "content-type": "application/json",
            "authorization": f"Bearer {access_token}",
        }

        response = requests.post(url, json=payload, headers=headers)
        print(response.text)
        print(headers)

        if response.status_code == 405:
            ctypes.windll.user32.MessageBoxW(
                0, "Forneça um link para a lousa!", "Miro API", 1
            )

        card_id = response.json()["id"]
        card_ids[card] = card_id

        x += 200
        if x > 600:
            y += 300
            x = 0


def create_tags():
    tag_ids = {}
    tag_colors = [
        "red",
        "light_green",
        "cyan",
        "yellow",
        "magenta",
        "green",
        "blue",
        "gray",
        "violet",
        "dark_green",
        "dark_blue",
        "black",
    ]

    for i, j in enumerate(set(card_tags.values())):
        for tag in j.split(","):
            url = f"https://api.miro.com/v2/boards/{board_id}/tags"

            payload = {"fillColor": f"{tag_colors[i]}", "title": f"{tag}"}
            headers = {
                "accept": "application/json",
                "content-type": "application/json",
                "authorization": f"Bearer {access_token}",
            }

            response = requests.post(url, json=payload, headers=headers)
            print(response.text)
            try:
                tag_ids[tag] = response.json()["id"]
            except KeyError:
                pass

    return tag_ids


def attach_tag_to_card(card_ids, tag_ids):
    for card_id in card_ids:
        card_name = card_ids[card_id]
        if (
            card_name in card_tags
        ):  # se o nome nao estiver nas tag, significa que esse post it ja estava na lousa
            # entao o codigo nao deve anexar nenhuma tag nele
            for tag_name_for_this_card in card_tags[card_name].split(","):
                if tag_name_for_this_card == "":
                    continue

                tag_id = tag_ids[tag_name_for_this_card]

                url = f"https://api.miro.com/v2/boards/{board_id}/items/{card_id}?tag_id={tag_id}"
                print(url)

                headers = {
                    "accept": "application/json",
                    "authorization": f"Bearer {access_token}",
                }

                response = requests.post(url, headers=headers)
                print(response)

    ctypes.windll.user32.MessageBoxW(
        0, "Dados importados!\n(atualize a página para ver as tags)", "Miro API", 1
    )


# ----- puxando dados do Miro e salvando na base
def remove_html_tag_from_strig(str):
    # quando tem quebra de linha ele gera um html que faz o card_name ficar todo feio (cheio de tags html)
    # tem jeitos melhores de dar esse replace, mas quis importar o minimo de bibliotecas possivel para manter o exe mais leve
    # deveria usar regex aqui. E espero que nao tenha outra tag que nao seja p e br

    return str.replace(r"<p>", "").replace(r"</p>", "").replace(r"<br>", "\n")


def get_updated_tags():
    tag_ids = get_tags_on_board()

    updated_card_tags = []
    for tag_name in tag_ids:
        tag_id = tag_ids[tag_name]

        url = (
            f"https://api.miro.com/v2/boards/{board_id}/items?limit=50&tag_id={tag_id}"
        )

        headers = {
            "accept": "application/json",
            "authorization": f"Bearer {access_token}",
        }

        response = requests.get(url, headers=headers)
        print(response.json())

        for card in response.json()["data"]:
            if "data" in card:
                card_name = card["data"]["content"]
                card_name = remove_html_tag_from_strig(card_name)
                updated_card_tags.append(
                    {"Dados": card_name, "Classificacoes": tag_name}
                )  # desse jeito ja ta pronto pra se mudar pra mais de uma classficação

    global df_new_classification
    df_new_classification = pd.DataFrame(updated_card_tags)

    # juntando as várias tags que um mesmo card pode ter em uma linha só:
    df_new_classification = (
        df_new_classification.groupby("Dados")["Classificacoes"]
        .apply(lambda x: ",".join(np.unique(x)))
        .reset_index()
    )

    # adicionando cards sem tags:
    cards = get_cards_on_board()
    cards = cards.values()
    print(cards)
    print(df_new_classification["Dados"].unique())
    cards = [remove_html_tag_from_strig(i) for i in cards]  # tirando tag html
    cards = [i for i in cards if i not in df_new_classification["Dados"].unique()]
    for card in cards:
        df_new_classification = pd.concat(
            [
                df_new_classification,
                pd.DataFrame([{"Dados": card, "Classificacoes": ""}]),
            ]
        )

    # salvando excel:
    try:
        df_new_classification.to_excel("Nova classificação.xlsx", index=False)
    except PermissionError:
        ctypes.windll.user32.MessageBoxW(
            0,
            "Feche o arquivo 'Nova classificação.xlsx' e clique novamente em exportar.",
            "Miro API",
            1,
        )
    except:
        ctypes.windll.user32.MessageBoxW(
            0, "Algo deu errado ao tentar salvar a base de dados", "Miro API", 1
        )

    # criando log de modificação:
    try:  # coloquei num try pq isso nao é fundamental pro funcionamento do programa, se der erro só printa o erro e segue a vida
        df_log = database.merge(
            df_new_classification,
            how="outer",
            on="Dados",
            suffixes=["_antigas", "_novas"],
        )
        df_log = df_log[
            df_log["Classificacoes_novas"] != df_log["Classificacoes_antigas"]
        ]
        df_log["Change_Time"] = datetime.datetime.now()
        df_log.to_excel(
            f"log_modificacao_{datetime.datetime.now().strftime('%Y-%m-%d, %H-%M-%S')}.xlsx"
        )  # se quiser que salva mais de um por dia é só trocar o argumento do strftime pra incluir horas minutos e segundos
    except:
        ctypes.windll.user32.MessageBoxW(
            0, "Nao foi possivel salvar logs de modificação", "Miro API", 1
        )

    ctypes.windll.user32.MessageBoxW(0, "Tags exportadas", "Miro API", 1)

    return df_new_classification


# ---- main -----

# le o token de acesso ao ambiente da lousa, se nao achar ele pede para inputar
try:
    access_token = pd.read_csv(
        "C:\\Users\\mgodoi\\OneDrive - M Square Investimentos Ltda\\Desktop\\TCC\\Final\\credentials.csv"
    ).iloc[0, 0]
except:
    ctypes.windll.user32.MessageBoxW(
        0,
        "Credenciais nao encontradas, por favor siga as instruções do botão de ajuda para conectar-se a API",
        "Miro API",
        1,
    )


board_id = ""

# ----- escolha da base de dados

try:
    excel_path = pd.read_csv("last_database.csv").iloc[0, 0]
except:
    ctypes.windll.user32.MessageBoxW(
        0,
        "Não foi possível conectar a base de dados da ultima seção, favor escolher um arquivo",
        "Miro API",
        1,
    )
card_tags, database = read_datase(excel_path, from_credentials=True)

print(card_tags)


root = Tk()
root.title("Miro API")
root.geometry("460x300")

frm = ttk.Frame(root, padding=10)
frm.place(x=0, y=0, width=460, height=300)

ttk.Label(frm, text="Insira o link da lousa abaixo:").place(x=0, y=0)

board_link = Text(frm, height=2, width=50)
board_link.place(x=20, y=30)
ttk.Button(
    frm,
    text="Conectar",
    width=21,
    padding=5,
    command=lambda: retrieve_input(board_link),
).place(x=20, y=80)

ttk.Button(
    frm,
    text="Escolher base de dados",
    padding=5,
    width=21,
    command=lambda: show_select_database_dialog(),
).place(x=285, y=80)

ttk.Button(
    frm,
    text="Importar Cards",
    padding=5,
    width=21,
    command=lambda: post_card_and_tags(),
).place(x=20, y=150)

ttk.Button(
    frm, text="Exportar Cards", padding=5, width=21, command=lambda: get_updated_tags()
).place(x=285, y=150)

ttk.Button(
    frm,
    text="Instruções",
    padding=5,
    width=15,
    command=lambda: show_instrucoes_de_uso(),
).place(x=20, y=240)

ttk.Button(
    frm,
    text="Alterar token",
    padding=5,
    width=15,
    command=lambda: show_access_token_input(),
).place(x=150, y=240)


ttk.Button(
    frm, text="Ajuda", padding=5, width=15, command=lambda: show_enable_board_tutorial()
).place(x=305, y=240)


root.mainloop()
