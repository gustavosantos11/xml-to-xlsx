# type: ignore
import os
import xmltodict
import pandas as pd


def get_info(archive, values):
    with open(f"nfs/{archive}", "rb") as archive_xml:
        dic_archive = xmltodict.parse(archive_xml)

        if "NFe" in dic_archive:
            info_nf = dic_archive["NFe"]['infNFe']
        else:
            info_nf = dic_archive["nfeProc"]["NFe"]['infNFe']
        note_number = info_nf["@Id"]
        emissor = info_nf["emit"]['xNome']
        client_name = info_nf["dest"]['xNome']
        address = info_nf["dest"]['enderDest']
        if "vol" in info_nf["transp"]:
            weight = info_nf["transp"]['vol']['pesoB']
        else:
            weight = "Não Informado"
        values.append([note_number, emissor, client_name, address, weight])


archives = os.listdir("nfs")

columns = ["numero_nota", "emissor", "nome_do_cliente", "endereco", "peso"]
values = []

for archive in archives:
    get_info(archive, values)


table = pd.DataFrame(columns=columns, data=values)
table.to_excel("3NotasFiscais.xlsx", index=False)
