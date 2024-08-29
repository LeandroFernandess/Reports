"""

Filename: FormatStockUDC.py
Developer: Leandro Fernandes
Date: 17/08/2024
Description: O código formata um arquivo excel e cria vários outros com base nos critérios estabelecidos pelas funções.

Chaves para pesquisa:

- Função para formatar o arquivo contendo os itens em box ==> FormatBox

- Função para formatar o arquivo contendo os itens em I2 ==> FormatI2

- Função para formatar o arquivo contendo os itens alocados em locações virtuais ==> FormatVirtualLoc

- Função para formatar o arquivo contendo os itens com missões pendentes de confirmação ==> FormatMissions

- Função para formatar o arquivo contendo os itens com pendência de armazenamento ==> FormatStorage

- Função para formatar o arquivo contendo os itens travados nas locações de qualidade ==> FormatQuality

- Função para formatar o arquivo contendo os itens que foram feitos retorno de linha ==> FormatRTL

- Função para formatar o arquivo contendo os itens que estão alocados em locação de transferência entre armazéns (Shuttle_load_CL) ==> FormatSHTLoad

- Função para formatar o arquivo contendo os itens que estão alocados em locação de transferência entre armazéns (Shuttle_trs_PC) ==> FormatTRS

- Função para formatar o arquivo contendo os itens que estão fora do padrão estabelecido (Fora de UDC'S container) ==> FormatContainer

- Função para formatar o arquivo contendo os itens rejeitados ==> FormatRej


"""

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

import sys
import os
import pandas as pd
import numpy as np
from datetime import datetime
from .Function import (
    RemoveColumnsBox,
    RemoveColumnsI2,
    RemoveColumnsVirtualLoc,
    RemoveColumnsMissions,
    RemoveColumnsStorage,
    RemoveColumnsQuality,
    RemoveColumnsRTL,
    RemoveColumnsSHTLoad,
    RemoveColumnsTRS,
    RemoveColumnsContainer,
    RemoveColumnsRej,
)
from datetime import timedelta
import locale

# Caminho para a pasta 'Relatórios'
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from Dictionaries.Dicts import (
    National,
    Foreign,
    External_shed,
    Responsibility,
    Users,
    Box,
    FirstT,
    SecondT,
    Destination,
)

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


def FormatBox():
    print("A formatação do arquivo contendo os materiais em box será iniciada!")
    Arquivo_Deem = pd.read_excel(
        r"C:\Users\F89074d\Desktop\Analises\Arquivos 2024\Divergências 2024.xlsx"
    )
    Data_base = pd.read_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Arquivos manuais\MM60.xlsx"
    )  # Modificar caminhos

    Data_frame = RemoveColumnsBox(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Estoque+udc.csv"
    )
    # Definindo "Material" como índice e a "Preço Unitário" como valor no dicionário:
    material_dict = Data_base.set_index("Material")["Preço"].to_dict()

    Data_frame["Data Carga"] = pd.to_datetime(
        Data_frame["Data Carga"], format="%d/%m/%y"
    )

    # Obtendo a data atual como Timestamp:
    today = pd.Timestamp(datetime.now().date())

    # Calculando a diferença em dias entre a data de hoje e a coluna 'Data Carga':
    Data_frame["Dias_Diferença"] = (today - Data_frame["Data Carga"]).dt.days

    # Definindo a função para verificar se o prazo está dentro ou fora do limite:
    def VerifyDays(dias):
        if dias > 5:
            return "Material fora do prazo de recebimento"
        else:
            return "Material dentro do prazo de recebimento"

    Data_frame["Prazo"] = Data_frame["Dias_Diferença"].apply(VerifyDays)

    Data_frame["Origem"] = None
    Data_frame["Origem"] = Data_frame["Código Fornecedor"].map(External_shed)
    Data_frame.loc[Data_frame["Razão Social Fornecedor"].isin(National), "Origem"] = (
        "Nacional"
    )
    Data_frame.loc[Data_frame["Razão Social Fornecedor"].isin(Foreign), "Origem"] = (
        "Importado"
    )

    Data_frame = Data_frame[
        (Data_frame["Locação"] == "I2") & (Data_frame["UdC Tipo"] == "BOX")
    ]
    # Configurando a coluna item para numérico:
    Data_frame["Item"] = Data_frame["Item"].apply(
        lambda x: pd.to_numeric(x, errors="coerce") if x.isdigit() else x
    )

    # Criando uma nova coluna 'Valor Unitário' com base no dicionário material_dict:
    Data_frame["Valor Unitário"] = Data_frame["Item"].map(material_dict)

    # Criar formatando a coluna 'Estocada' com os dígitos antes da vírgula, removendo os zeros extras:
    Data_frame["Estocada"] = Data_frame["Estocada"].apply(
        lambda x: int("".join(filter(str.isdigit, str(x)))[:-3])
    )

    Data_frame["Valor Total"] = Data_frame["Estocada"] * Data_frame["Valor Unitário"]

    # Excluindo a coluna "Dias Diferença":
    Data_frame = Data_frame.drop(["Dias_Diferença", "Locação"], axis=1)
    Data_frame.rename(columns={"Item": "PN", "RC": "Nº Viagem"}, inplace=True)

    # Mesclar os DataFrames com base nas colunas "PN" e "Nº viagem":
    Data_frame = Data_frame.merge(
        Arquivo_Deem[["PN", "Nº Viagem", "Comentário"]],
        how="left",
        on=["PN", "Nº Viagem"],
    )

    # Adicionando a verificação na coluna "Comentário":
    Data_frame["Comentário"] = np.where(
        Data_frame[["PN", "Nº Viagem"]]
        .isin(Arquivo_Deem[["PN", "Nº Viagem"]].to_dict(orient="list"))
        .all(axis=1),
        "É divergência",
        "Não é divergência",
    )

    # Remover duplicatas baseadas na coluna "Código UdC"
    Data_frame = Data_frame.drop_duplicates(subset="Código UdC")

    Data_frame = Data_frame[
        [
            "Código UdC",
            "PN",
            "UdC Tipo",
            "Estocada",
            "Valor Unitário",
            "Valor Total",
            "Nº Viagem",
            "Nota Fiscal",
            "Código Fornecedor",
            "Razão Social Fornecedor",
            "Origem",
            "Comentário",
            "Data Carga",
            "Prazo",
        ]
    ]

    Data_frame.to_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Formatados\Box CL.xlsx",
        index=False,
    )
    print(
        "\n",
        "A formatação do arquivo contendo os materiais em box foi finalizada com sucesso!",
        "\n",
    )


# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


def FormatI2():
    print(
        "A formatação do arquivo contendo os itens recebidos e alocados em I2 será iniciada!"
    )
    Data_base = pd.read_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Arquivos manuais\MM60.xlsx"
    )

    # Definindo "Material" como índice e a "Preço Unitário" como valor no dicionário:
    material_dict = Data_base.set_index("Material")["Preço"].to_dict()
    material_type = Data_base.set_index("Material")["TpM"].to_dict()

    Data_frame = RemoveColumnsI2(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Estoque+udc.csv"
    )
    Data_frame = Data_frame[Data_frame["Locação"] == "I2"]

    # Convertendo a coluna 'Update' para datetime:
    Data_frame["Update"] = pd.to_datetime(Data_frame["Update"], format="%d/%m/%y %H:%M")
    Data_frame["Data Carga"] = pd.to_datetime(
        Data_frame["Data Carga"], format="%d/%m/%y"
    )

    today = datetime.now().date()
    if today.weekday() == 0:  # 0 representa segunda feira
        # Filtrando o DataFrame para excluir as informações do dia atual e de 48 horas
        Data_frame = Data_frame[
            (Data_frame["Update"].dt.date != today)
            & (Data_frame["Update"].dt.date != (today - timedelta(days=2)))
        ]
    else:
        Data_frame = Data_frame[(Data_frame["Update"].dt.date != today)]

    list = ["BOX", "CAIXA"]
    Data_frame = Data_frame[~(Data_frame["UdC Tipo"].isin(list))]

    Data_frame["Embalagem"] = None
    Data_frame.loc[Data_frame["UdC Tipo"].isin(Box), "Embalagem"] = "Caixaria"
    Data_frame.loc[Data_frame["Embalagem"].isna(), "Embalagem"] = "Pesados"

    Data_frame["Localidade"] = None
    Data_frame["Localidade"] = Data_frame["Código Fornecedor"].map(External_shed)
    Data_frame.loc[
        Data_frame["Razão Social Fornecedor"].isin(National), "Localidade"
    ] = "Nacional"
    Data_frame.loc[
        Data_frame["Razão Social Fornecedor"].isin(Foreign), "Localidade"
    ] = "Importado"
    Data_frame.loc[Data_frame["Nota Fiscal"].isna(), "Localidade"] = "Analise a Parte"
    Data_frame.loc[Data_frame["Código Fornecedor"] == "BNH57", "Localidade"] = (
        "CNH INDUSTRIAL BRASIL LTDA"
    )

    Data_frame.loc[Data_frame["Origem"].isna(), "Origem"] = "Recebimento em I2"

    Data_frame["Nome do Colaborador"] = None
    Data_frame["Responsabilidade"] = None
    Data_frame["Nome do Colaborador"] = Data_frame["Usuario Modificação"].map(Users)
    Data_frame["Responsabilidade"] = Data_frame["Nome do Colaborador"].map(
        Responsibility
    )

    # Criando uma nova coluna 'Valor Unitário' com base no dicionário material_dict:
    Data_frame["Item"] = Data_frame["Item"].apply(
        lambda x: pd.to_numeric(x, errors="coerce") if x.isdigit() else x
    )
    Data_frame["Valor Unitário"] = Data_frame["Item"].map(material_dict)
    Data_frame["Depósito"] = Data_frame["Item"].map(material_type)

    lista_padrão = ["ND", "PD"]
    lista_sw = ["ZV", "ZB"]

    Data_frame["Tipo"] = None
    Data_frame.loc[Data_frame["Depósito"].isin(lista_sw), "Tipo"] = "Material é SW"
    Data_frame.loc[Data_frame["Depósito"].isin(lista_padrão), "Tipo"] = (
        "Material não é SW"
    )
    # Removendo os zeros extras:
    Data_frame["Estocada"] = Data_frame["Estocada"].apply(
        lambda x: int("".join(filter(str.isdigit, str(x)))[:-3])
    )

    Data_frame["Valor Total"] = Data_frame["Estocada"] * Data_frame["Valor Unitário"]

    # Ordenando as colunas:
    Data_frame = Data_frame[
        [
            "Código UdC",
            "Item",
            "Estocada",
            "UdC Tipo",
            "Embalagem",
            "Valor Unitário",
            "Valor Total",
            "Status Contábil",
            "Locação",
            "Origem",
            "Depósito",
            "Tipo",
            "RC",
            "Nota Fiscal",
            "Código Fornecedor",
            "Razão Social Fornecedor",
            "Localidade",
            "Usuario Modificação",
            "Nome do Colaborador",
            "Responsabilidade",
            "Data Carga",
            "Update",
        ]
    ]
    Data_frame.to_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Formatados\I2.xlsx",
        index=False,
    )
    print(
        "A formatação do arquivo contendo os itens recebidos e alocados em I2 foi finalizada com sucesso!",
        "\n",
    )


# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


def FormatVirtualLoc():
    print(
        "A formatação do arquivo contendo os itens alocados em locações virtuais será iniciada!"
    )
    Data_base = pd.read_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Arquivos manuais\MM60.xlsx"
    )
    Data_frame = RemoveColumnsVirtualLoc(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Estoque+udc.csv"
    )
    locations = [
        "CLP1.C.01.17.F.00",
        "CLP1.A.00.02.A.00",
        "WL38CL",
        "GR21",
        "SEPARACAO",
        "B3",
        "CLP1.R.02.01.A.00",
        "EXPT",
        "TRV2",
    ]
    Data_frame = Data_frame[(Data_frame["Locação"].isin(locations))]

    # Definindo "Material" como índice e a "Preço Unitário" como valor no dicionário:
    material_dict = Data_base.set_index("Material")["Preço"].to_dict()

    # Convertendo a coluna 'Data Carga' para datetime
    Data_frame["Update"] = pd.to_datetime(Data_frame["Update"], format="%d/%m/%y %H:%M")
    Data_frame["Data Carga"] = pd.to_datetime(Data_frame["Update"], format="%d/%m/%y")

    # Definindo o locale para português brasileiro
    locale.setlocale(locale.LC_TIME, "pt_BR")

    # Criando novas colunas:

    Data_frame["Nome do Colaborador"] = Data_frame["Usuario Modificação"].map(Users)
    Data_frame["Responsabilidade"] = Data_frame["Nome do Colaborador"].map(
        Responsibility
    )

    # Configurando a coluna item para numérico:
    Data_frame["Item"] = Data_frame["Item"].apply(
        lambda x: pd.to_numeric(x, errors="coerce") if x.isdigit() else x
    )
    Data_frame["Valor Unitário"] = Data_frame["Item"].map(material_dict)

    # Criar formatando a coluna "Estocada" com os dígitos antes da vírgula, removendo os zeros extras:
    Data_frame["Estocada"] = Data_frame["Estocada"].apply(
        lambda x: int("".join(filter(str.isdigit, str(x)))[:-3])
    )
    Data_frame["Valor Total"] = Data_frame["Estocada"] * Data_frame["Valor Unitário"]
    # Ordenando as colunas:
    Data_frame = Data_frame[
        [
            "Código UdC",
            "Item",
            "UdC Tipo",
            "Estocada",
            "Valor Unitário",
            "Valor Total",
            "Locação",
            "Origem",
            "Destino",
            "Usuario Modificação",
            "Nome do Colaborador",
            "Responsabilidade",
            "Data Carga",
            "Update",
        ]
    ]
    Data_frame.to_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Formatados\Locações Virtuais.xlsx",
        index=False,
    )

    print(
        "A formatação do arquivo contendo os itens alocados em locações virtuais foi finalizada com sucesso!",
        "\n",
    )


# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


def FormatMissions():
    print(
        "A formatação do arquivo contendo os itens com pendência de confirmação de missões será iniciada!"
    )
    Data_frame = RemoveColumnsMissions(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Estoque+udc.csv"
    )
    Data_base = pd.read_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Arquivos manuais\MM60.xlsx"
    )
    material_dict = Data_base.set_index("Material")["Preço"].to_dict()

    Data_frame = Data_frame[Data_frame["Site"] == "CENTRO_LOGISTICO"]

    # Excluindo os valores nulos da coluna desejada:
    Data_frame = Data_frame.dropna(subset=["Destino"])
    today = datetime.now().strftime("%d/%m/%y")
    # Criando uma nova coluna com apenas as informações de Dia, Mês e Ano:
    Data_frame["Data de criação da missão"] = Data_frame["Update"].str.slice(stop=8)
    Data_frame["Data de criação da missão"] = pd.to_datetime(
        Data_frame["Data de criação da missão"], format="%d/%m/%y"
    )
    # Criando a coluna com o valor das horas e minutos:
    Data_frame["Horário de criação da missão"] = Data_frame["Update"].str.slice(
        start=-6
    )
    # Retirando os espaços em branco:
    Data_frame["Horário de criação da missão"] = Data_frame[
        "Horário de criação da missão"
    ].str.strip()

    Data_frame["Horário de criação da missão"] = pd.to_datetime(
        Data_frame["Horário de criação da missão"], format="%H:%M"
    )
    # Formatando o horário para apenas hora e minuto
    Data_frame["Horário de criação da missão"] = Data_frame[
        "Horário de criação da missão"
    ].dt.strftime("%H:%M")

    Data_frame["Responsabilidade"] = np.where(
        (
            (Data_frame["Horário de criação da missão"] >= "05:00")
            & (Data_frame["Horário de criação da missão"] < "16:00")
        ),
        "Missão criada no 1° turno",
        np.where(
            (
                (Data_frame["Horário de criação da missão"] >= "16:00")
                | (Data_frame["Horário de criação da missão"] < "02:00")
            ),
            "Missão criada no 2° turno",
            None,  # Valor padrão se nenhuma das condições acima for atendida
        ),
    )

    Data_frame = Data_frame[Data_frame["Data de criação da missão"] != today]
    # Formatando a coluna "Estocada" com os dígitos antes da vírgula, removendo os zeros extras:
    Data_frame["Estocada"] = Data_frame["Estocada"].apply(
        lambda x: int("".join(filter(str.isdigit, str(x)))[:-3])
    )
    # Configurando a coluna "Item" para numérico:
    Data_frame["Item"] = Data_frame["Item"].apply(
        lambda x: pd.to_numeric(x, errors="coerce") if x.isdigit() else x
    )

    Data_frame["Valor Unitário"] = Data_frame["Item"].map(material_dict)

    Data_frame["Valor Total"] = Data_frame["Estocada"] * Data_frame["Valor Unitário"]
    # Ordenando as colunas:
    Data_frame = Data_frame[
        [
            "Código UdC",
            "Item",
            "UdC Tipo",
            "UdC Container",
            "Status Contábil",
            "Estocada",
            "Valor Unitário",
            "Valor Total",
            "Locação",
            "Destino",
            "Data de criação da missão",
            "Horário de criação da missão",
            "Responsabilidade",
        ]
    ]

    Data_frame.to_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Formatados\Missões Pendentes.xlsx",
        index=False,
    )
    print(
        "A formatação do arquivo contendo os itens com pendência de confirmação de missões foi finalizada com sucesso!",
        "\n",
    )


# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


def FormatStorage():
    print(
        "A formatação do arquivo contendo os itens pendentes de armazenamento será iniciada!"
    )
    Data_frame = RemoveColumnsStorage(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Estoque+udc.csv"
    )
    Data_base = pd.read_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Arquivos manuais\MM60.xlsx"
    )
    # Definindo "Material" como índice e a "Preço Unitário" como valor no dicionário:
    material_dict = Data_base.set_index("Material")["Preço"].to_dict()

    locations = ["CLR1", "CLR2", "I2"]
    Data_frame = Data_frame[
        (Data_frame["Locação"].isin(locations)) & (Data_frame["UdC Tipo"] != "BOX")
    ]

    Data_frame["Update"] = pd.to_datetime(
        Data_frame["Update"], format="%d/%m/%y %H:%M", errors="coerce"
    )

    today = datetime.now()

    # Calculando a diferença entre a coluna 'Update' e a data/hora atual:
    Data_frame["Diferença"] = today - Data_frame["Update"]

    Data_frame["Prazo"] = None
    Data_frame.loc[Data_frame["Diferença"] > timedelta(hours=48), "Prazo"] = (
        "Fora do prazo de armazenamento"
    )
    Data_frame.loc[Data_frame["Diferença"] <= timedelta(hours=48), "Prazo"] = (
        "Dentro do prazo de armazenamento"
    )
    Data_frame = Data_frame.drop("Diferença", axis=1)

    # Definindo o locale para português brasileiro
    locale.setlocale(locale.LC_TIME, "pt_BR")

    # Configurando a coluna item para numérico:
    Data_frame["Item"] = Data_frame["Item"].apply(
        lambda x: pd.to_numeric(x, errors="coerce") if x.isdigit() else x
    )

    # Criando uma nova coluna 'Valor Unitário' com base no dicionário material_dict:
    Data_frame["Valor Unitário"] = Data_frame["Item"].map(material_dict)

    # Criar ajustando a coluna "Estocada" com os dígitos antes da vírgula, removendo os zeros extras:
    Data_frame["Estocada"] = Data_frame["Estocada"].apply(
        lambda x: int("".join(filter(str.isdigit, str(x)))[:-3])
    )

    Data_frame["Valor Total"] = Data_frame["Estocada"] * Data_frame["Valor Unitário"]

    Data_frame["Nome do Colaborador"] = None
    Data_frame["Responsabilidade"] = None
    Data_frame["Nome do Colaborador"] = Data_frame["Usuario Modificação"].map(Users)
    Data_frame["Responsabilidade"] = Data_frame["Nome do Colaborador"].map(
        Responsibility
    )

    Data_frame["Turno"] = None
    Data_frame.loc[Data_frame["Responsabilidade"].isin(FirstT), "Turno"] = (
        "Primeiro Turno"
    )
    Data_frame.loc[Data_frame["Responsabilidade"].isin(SecondT), "Turno"] = (
        "Segundo Turno"
    )

    # Colocando em ordem as colunas:
    Data_frame = Data_frame[
        [
            "Código UdC",
            "Item",
            "UdC Tipo",
            "Locação",
            "Estocada",
            "Valor Unitário",
            "Valor Total",
            "Status Contábil",
            "RC",
            "Nota Fiscal",
            "Código Fornecedor",
            "Razão Social Fornecedor",
            "Usuario Modificação",
            "Nome do Colaborador",
            "Responsabilidade",
            "Turno",
            "Prazo",
            "Inserção",
            "Update",
        ]
    ]
    Data_frame.to_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Formatados\Pendência armazenagem.xlsx",
        index=False,
    )
    print(
        "A formatação do arquivo contendo os itens pendentes de armazenamento foi finalizada com sucesso!",
        "\n",
    )


# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


def FormatQuality():
    print(
        "A formatação do arquivo contendo os itens com status contábil de Qualidade será iniciada!"
    )
    Data_frame = RemoveColumnsQuality(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Estoque+udc.csv"
    )
    Data_base = pd.read_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Arquivos manuais\MM60.xlsx"
    )
    # Definindo "Material" como índice e a "Preço Unitário" como valor no dicionário:
    material_dict = Data_base.set_index("Material")["Preço"].to_dict()

    Data_frame = Data_frame[Data_frame["Site"] == "CENTRO_LOGISTICO"]

    Data_frame["Data Carga"] = pd.to_datetime(
        Data_frame["Data Carga"], format="%d/%m/%y"
    )
    Data_frame["Update"] = pd.to_datetime(Data_frame["Update"], format="%d/%m/%y %H:%M")

    today = datetime.now()
    if today.weekday() == 0:  # 0 representa segunda feira
        # Filtrando o DataFrame para excluir as informações do dia atual e do dia anterior:
        Data_frame = Data_frame[
            (Data_frame["Data Carga"].dt.date != today)
            & (Data_frame["Data Carga"].dt.date != (today - timedelta(days=3)))
        ]
    else:
        Data_frame = Data_frame[(Data_frame["Data Carga"].dt.date != today)]

    # Configurando a coluna item para numérico:
    Data_frame["Item"] = Data_frame["Item"].apply(
        lambda x: pd.to_numeric(x, errors="coerce") if x.isdigit() else x
    )
    Data_frame["Valor Unitário"] = Data_frame["Item"].map(material_dict)

    # Formatando a coluna "Estocada" com os dígitos antes da vírgula, removendo os zeros extras:
    Data_frame["Estocada"] = Data_frame["Estocada"].apply(
        lambda x: int("".join(filter(str.isdigit, str(x)))[:-3])
    )
    Data_frame["Valor Total"] = Data_frame["Estocada"] * Data_frame["Valor Unitário"]

    Data_frame["Prazo"] = None
    prazo_48h = timedelta(hours=48)

    # Verificando se a diferença entre a data atual e 'Datas Movimento' é menor ou igual a 48 horas
    Data_frame.loc[(today - Data_frame["Update"]) <= prazo_48h, "Prazo"] = (
        "Movimentação sistêmica dentro de 48 horas"
    )
    Data_frame.loc[Data_frame["Prazo"].isna(), "Prazo"] = (
        "Movimentação sistêmica anterior a 48 horas"
    )

    status = ["Retorno para Forn", "Qualidade", "Scrap"]
    Data_frame = Data_frame[(Data_frame["Status Contábil"].isin(status))]

    Data_frame = Data_frame[Data_frame["Locação"] != "U02"]

    Data_frame = Data_frame[
        [
            "Código UdC",
            "Item",
            "UdC Tipo",
            "Status Contábil",
            "Estocada",
            "Valor Unitário",
            "Valor Total",
            "Locação",
            "RC",
            "Código Fornecedor",
            "Razão Social Fornecedor",
            "Data Carga",
            "Update",
            "Prazo",
        ]
    ]
    Data_frame.to_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Formatados\Itens em qualidade.xlsx",
        index=False,
    )
    print(
        "A formatação do arquivo contendo os itens com status contábil de qualidade foi finalizada com sucesso!",
        "\n",
    )


# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


def FormatRTL():
    print(
        "A formatação do arquivo contendo os itens que foram feitos o movimento de RTL será iniciada!"
    )
    Data_frame = RemoveColumnsRTL(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Estoque+udc.csv"
    )

    Data_frame = Data_frame[Data_frame["Locação"] == "RTL2"]

    Data_frame["Nome do Colaborador"] = Data_frame["Usuario Modificação"].map(Users)
    Data_frame["Condutor"] = Data_frame["Nome do Colaborador"].map(Responsibility)

    Data_frame["Update"] = pd.to_datetime(
        Data_frame["Update"], format="%d/%m/%y %H:%M", errors="coerce"
    )

    today = datetime.now()

    # Calculando a diferença entre a coluna 'Update' e a data/hora atual
    Data_frame["Diferença"] = today - Data_frame["Update"]

    Data_frame["Prazo"] = None
    Data_frame.loc[Data_frame["Diferença"] > timedelta(hours=24), "Prazo"] = (
        "RTL 2 anterior a 24 horas"
    )
    Data_frame.loc[Data_frame["Diferença"] <= timedelta(hours=24), "Prazo"] = (
        "RTL2 dentro de 24 horas"
    )
    # Definindo o locale para português brasileiro
    locale.setlocale(locale.LC_TIME, "pt_BR")

    Data_frame = Data_frame.drop("Diferença", axis=1)
    Data_frame.to_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Formatados\Itens em RTL2.xlsx",
        index=False,
    )
    print(
        "A formatação do arquivo contendo os itens que foram feitos o movimento de RTL foi finalizada com sucesso!",
        "\n",
    )


# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


def FormatSHTLoad():
    print(
        "A formatação do arquivo contendo os itens alocados em Shuttle_load_cl será iniciada!"
    )
    # Definindo o locale para português brasileiro
    locale.setlocale(locale.LC_TIME, "pt_BR")

    Data_frame = RemoveColumnsSHTLoad(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Estoque+udc.csv"
    )
    Data_base = pd.read_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Arquivos manuais\MM60.xlsx"
    )

    # Definindo 'Material' como índice e convertendo para dicionário
    tipo_material = Data_base.set_index("Material")["TpM"].to_dict()
    preço_material = Data_base.set_index("Material")["Preço"].to_dict()

    Data_frame = Data_frame[Data_frame["Locação"] == "SHUTTLE_LOAD_CL"]

    Data_frame["Update"] = pd.to_datetime(Data_frame["Update"], format="%d/%m/%y %H:%M")
    Data_frame["Data Carga"] = pd.to_datetime(
        Data_frame["Data Carga"], format="%d/%m/%y"
    )

    # Tratando os dados da coluna "Código Produto":
    Data_frame["Código Produto"] = Data_frame["Item"].apply(
        lambda x: pd.to_numeric(x, errors="coerce") if x.isdigit() else x
    )

    Data_frame["Tipo"] = Data_frame["Código Produto"].map(tipo_material)

    today = datetime.now()

    Data_frame["Prazo"] = None
    prazo_24h = timedelta(hours=24)

    # Verificando se a diferença entre a data atual e 'Update' é menor ou igual a 24 horas
    Data_frame.loc[(today - Data_frame["Update"]) <= prazo_24h, "Prazo"] = (
        "Separação dentro de 24 horas"
    )
    # Preenchendo 'Fora do prazo de 24 horas' onde 'Prazo' ainda não está definido
    Data_frame.loc[Data_frame["Prazo"].isna(), "Prazo"] = (
        "Separação Anterior a 24 horas"
    )

    lista_padrão = ["ND", "PD"]
    lista_sw = ["ZV", "ZB"]

    Data_frame["Depósito"] = None
    Data_frame.loc[Data_frame["Tipo"].isin(lista_padrão), "Depósito"] = (
        "Material não é SW"
    )
    Data_frame.loc[Data_frame["Tipo"].isin(lista_sw), "Depósito"] = "Material é SW"

    condition = [
        (Data_frame["Update"].dt.hour >= 6) & (Data_frame["Update"].dt.hour < 16),
        (Data_frame["Update"].dt.hour >= 16)
        | (
            Data_frame["Update"].dt.hour < 2
        ),  # "|" é o operador lógico "OU" para combinar condições
    ]

    # Defina os valores correspondentes

    work_shift = ["Primeiro Turno", "Segundo Turno"]
    # Crie a nova coluna com base nas condições
    Data_frame["Turno"] = np.select(condition, work_shift, default="Primeiro Turno")

    # Configurando a coluna item para numérico:
    Data_frame["Item"] = Data_frame["Item"].apply(
        lambda x: pd.to_numeric(x, errors="coerce") if x.isdigit() else x
    )
    Data_frame["Valor Unitário"] = Data_frame["Item"].map(preço_material)

    # Criar uma nova coluna 'Estocada_Sem_Virgula' com os dígitos antes da vírgula, removendo os zeros extras
    Data_frame["Estocada"] = Data_frame["Estocada"].apply(
        lambda x: int("".join(filter(str.isdigit, str(x)))[:-3])
    )
    Data_frame["Valor Total"] = Data_frame["Estocada"] * Data_frame["Valor Unitário"]

    Data_frame = Data_frame[
        [
            "Código UdC",
            "Item",
            "Estocada",
            "UdC Tipo",
            "Valor Unitário",
            "Valor Total",
            "Tipo",
            "Depósito",
            "Locação",
            "RC",
            "Nota Fiscal",
            "Código Fornecedor",
            "Razão Social Fornecedor",
            "Usuario Modificação",
            "Turno",
            "Data Carga",
            "Update",
            "Prazo",
        ]
    ]
    Data_frame.to_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Formatados\Itens em Shuttle_Load_CL.xlsx",
        index=False,
    )
    print(
        "A formatação do arquivo contendo os itens alocados em Shuttle_load_cl foi finalizada com sucesso!",
        "\n",
    )


# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


def FormatTRS():
    print(
        "A formatação do arquivo contendo os itens que estão pendentes de recebimento shuttle_trs_pc será iniciada!"
    )
    Data_frame = RemoveColumnsTRS(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Estoque+udc.csv"
    )
    Data_base = pd.read_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Arquivos manuais\MM60.xlsx"
    )
    # Definindo "Material" como índice e a "Preço Unitário" como valor no dicionário:
    material_dict = Data_base.set_index("Material")["Preço"].to_dict()

    Data_frame = Data_frame[Data_frame["Locação"] == "SHUTTLE_TRS_PC"]

    Data_frame["Update"] = pd.to_datetime(Data_frame["Update"], format="%d/%m/%y %H:%M")
    Data_frame["Data Carga"] = pd.to_datetime(
        Data_frame["Data Carga"], format="%d/%m/%y"
    )

    today = datetime.now().date()
    if today.weekday() == 0:  # 0 representa segunda feira
        # Filtrando o DataFrame para excluir as informações do dia atual e do dia anterior
        Data_frame = Data_frame[
            (Data_frame["Update"].dt.date != today)
            & (Data_frame["Update"].dt.date != (today - timedelta(days=3)))
        ]
    else:
        Data_frame = Data_frame[(Data_frame["Update"].dt.date != today)]

    # Criando uma nova coluna 'Valor Unitário' com base no dicionário material_dict:
    Data_frame["Item"] = Data_frame["Item"].apply(
        lambda x: pd.to_numeric(str(x), errors="coerce") if str(x).isdigit() else x
    )
    Data_frame["Valor Unitário"] = Data_frame["Item"].map(material_dict)

    # Criar uma nova coluna 'Estocada_Sem_Virgula' com os dígitos antes da vírgula, removendo os zeros extras
    Data_frame["Estocada"] = Data_frame["Estocada"].apply(
        lambda x: int("".join(filter(str.isdigit, str(x)))[:-3])
    )
    Data_frame["Valor Total"] = Data_frame["Estocada"] * Data_frame["Valor Unitário"]
    # Defina as condições
    condition = [
        (Data_frame["Update"].dt.hour >= 6) & (Data_frame["Update"].dt.hour < 14),
        (Data_frame["Update"].dt.hour >= 22) | (Data_frame["Update"].dt.hour < 2),
        (Data_frame["Update"].dt.hour >= 14)
        | (
            Data_frame["Update"].dt.hour < 2
        ),  # "|" é o operador lógico "OU" para combinar condições
    ]

    work_shift = ["Primeiro Turno", "Primeiro Turno", "Segundo Turno"]
    # Crie a nova coluna com base nas condições
    Data_frame["Turno"] = np.select(condition, work_shift, default="Primeiro Turno")

    Data_frame.to_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Formatados\Itens em Shuttle_TRS_PC.xlsx",
        index=False,
    )

    print(
        "A formatação do arquivo contendo os itens que estão pendentes de recebimento shuttle_trs_pc foi finalizada com sucesso!",
        "\n",
    )


# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


def FormatContainer():
    print(
        "A formatação do arquivo contendo os itens fora de UDC container será iniciada!"
    )
    Data_frame = RemoveColumnsContainer(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Estoque+udc.csv"
    )

    Data_base = pd.read_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Arquivos manuais\MM60.xlsx"
    )

    # Definindo "Material" como índice e a "Preço Unitário" como valor no dicionário:
    material_dict = Data_base.set_index("Material")["Preço"].to_dict()
    material_type = Data_base.set_index("Material")["TpM"].to_dict()

    location = ["CLR1", "CLR2", "I2"]
    packaging = [
        "KLT1",
        "KLT2",
        "KLT3",
        "KLT4",
        "KLT5",
        "KLT6",
        "KLT7",
        "KLT8",
        "KLT9",
        "KLT10",
    ]

    Data_frame = Data_frame[
        Data_frame["Locação"].isin(location) & (Data_frame["UdC Tipo"].isin(packaging))
    ]

    Data_frame["Nome do Colaborador"] = Data_frame["Usuario Modificação"].map(Users)
    Data_frame["Local"] = Data_frame["Nome do Colaborador"].map(Destination)
    Data_frame["Responsável"] = Data_frame["Nome do Colaborador"].map(Responsibility)

    # Obtendo apenas os valores nulos da coluna "Udc Container":
    null = Data_frame[Data_frame["UdC Container"].isna()]
    Data_frame = null
    Data_frame = Data_frame.drop(["UdC Container"], axis=1)

    # Configurando a coluna item para numérico:
    Data_frame["Item"] = Data_frame["Item"].apply(
        lambda x: pd.to_numeric(x, errors="coerce") if x.isdigit() else x
    )
    Data_frame["Valor Unitário"] = Data_frame["Item"].map(material_dict)

    # Formatando a coluna "Estocada" com os dígitos antes da vírgula, removendo os zeros extras:
    Data_frame["Estocada"] = Data_frame["Estocada"].apply(
        lambda x: int("".join(filter(str.isdigit, str(x)))[:-3])
    )
    Data_frame["Valor Total"] = Data_frame["Estocada"] * Data_frame["Valor Unitário"]

    Data_frame["Update"] = pd.to_datetime(Data_frame["Update"], format="%d/%m/%y %H:%M")
    Data_frame["Data Carga"] = pd.to_datetime(
        Data_frame["Data Carga"], format="%d/%m/%y"
    )

    today = datetime.now().date()
    if today.weekday() == 0:  # 0 representa segunda feira
        # Filtrando o DataFrame para excluir as informações do dia atual e do dia anterior:
        Data_frame = Data_frame[
            (Data_frame["Update"].dt.date != today)
            & (Data_frame["Update"].dt.date != (today - timedelta(days=3)))
        ]  # Tratando a informação caso a data atual seja segunda feira
    else:
        Data_frame = Data_frame[(Data_frame["Update"].dt.date != today)]

    Data_frame["Depósito"] = Data_frame["Item"].map(material_type)

    lista_padrão = ["ND", "PD"]
    lista_sw = ["ZV", "ZB"]
    lista_embalagem = ["KLT1", "KLT2"]

    Data_frame["Tipo"] = None
    Data_frame.loc[Data_frame["Depósito"].isin(lista_padrão), "Tipo"] = (
        "Material não é SW"
    )
    Data_frame.loc[Data_frame["Depósito"].isin(lista_sw), "Tipo"] = "Material é SW"
    Data_frame.loc[Data_frame["UdC Tipo"].isin(lista_embalagem), "Tipo"] = (
        "Material é SW"
    )

    hour = datetime.now()
    Data_frame["Diferença"] = hour - Data_frame["Update"]

    Data_frame["Prazo"] = None
    Data_frame.loc[Data_frame["Diferença"] > timedelta(hours=24), "Prazo"] = (
        "Fora do prazo de alocação em UDC Container"
    )
    Data_frame.loc[Data_frame["Diferença"] <= timedelta(hours=24), "Prazo"] = (
        "Dentro do prazo de alocação em UDC Container"
    )
    Data_frame = Data_frame.drop("Diferença", axis=1)

    Data_frame = Data_frame[
        [
            "Código UdC",
            "Item",
            "Estocada",
            "Valor Unitário",
            "Valor Total",
            "UdC Tipo",
            "Locação",
            "Status Contábil",
            "Destino",
            "Depósito",
            "Tipo",
            "Usuario Modificação",
            "Nome do Colaborador",
            "Responsável",
            "Local",
            "Data Carga",
            "Update",
            "Prazo",
        ]
    ]

    Data_frame.to_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Formatados\Itens fora de container.xlsx",
        index=False,
    )
    print(
        "A formatação do arquivo contendo os itens fora de UDC container foi finalizada com sucesso!",
        "\n",
    )


# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


def FormatRej():
    print("A formatação do arquivo contendo os itens rejeitados será iniciada!")

    Data_frame = RemoveColumnsRej(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Estoque+udc.csv"
    )

    Data_base = pd.read_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Base\Arquivos manuais\MM60.xlsx"
    )

    material_dict = Data_base.set_index("Material")["Preço"].to_dict()
    material_desc = Data_base.set_index("Material")["Texto breve material"].to_dict()

    rejected_p1 = ["LRSI", "LRRE", "LRBT", "LRGP", "LRSQ"]
    rejected_p2 = ["LRPT"]

    repair_p1 = ["LCRE", "LCBT", "LCGP", "LCSI", "LCSQ"]
    repair_p2 = ["LCPT"]

    all_types = rejected_p1 + rejected_p2 + repair_p1 + repair_p2

    Data_frame = Data_frame[Data_frame["Locação"].isin(all_types)]

    map_type = {
        **{k: "Pátio 1" for k in rejected_p1},
        **{k: "Pátio 2" for k in rejected_p2},
        **{k: "Pátio 1" for k in repair_p1},
        **{k: "Pátio 2" for k in repair_p2},
    }
    Data_frame["Pátio"] = Data_frame["Locação"].map(map_type)

    Data_frame["Status validação"] = Data_frame["Locação"].map(
        lambda x: "Rejeitado" if x in rejected_p1 + rejected_p2 else "Conserto"
    )

    Data_frame["Update"] = pd.to_datetime(Data_frame["Update"], format="%d/%m/%y %H:%M")
    Data_frame["Data Carga"] = pd.to_datetime(
        Data_frame["Data Carga"], format="%d/%m/%y"
    )

    # Configurando a coluna item para numérico:
    Data_frame["Item"] = Data_frame["Item"].apply(
        lambda x: pd.to_numeric(x, errors="coerce") if x.isdigit() else x
    )
    Data_frame["Valor Unitário"] = Data_frame["Item"].map(material_dict)
    Data_frame["Descrição"] = Data_frame["Item"].map(material_desc)

    # Criar formatando a coluna 'Estocada' com os dígitos antes da vírgula, removendo os zeros extras:
    Data_frame["Estocada"] = Data_frame["Estocada"].apply(
        lambda x: int("".join(filter(str.isdigit, str(x)))[:-3])
    )

    Data_frame["Valor Total"] = Data_frame["Estocada"] * Data_frame["Valor Unitário"]

    Data_frame = Data_frame[
        [
            "Código UdC",
            "Item",
            "Descrição",
            "Estocada",
            "Valor Unitário",
            "Valor Total",
            "UdC Tipo",
            "UdC Container",
            "Status Contábil",
            "Locação",
            "Status validação",
            "Pátio",
            "Data Carga",
            "Update",
        ]
    ]

    Data_frame.to_excel(
        r"C:\Users\F89074d\Documents\Python - Arquivos Formatados\Rejeitado.xlsx",
        index=False,
    )

    print(
        "A formatação do arquivo contendo os itens rejeitados foi finalizado com sucesso!",
        "\n",
    )
