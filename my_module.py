# mymodule.py

import pandas as pd

def create_dataframe(data_list):
    """
    Tworzy obiekt DataFrame z biblioteki pandas na podstawie listy.
    :param data_list: Lista zawierająca dane do umieszczenia w DataFrame.
    :return: Obiekt DataFrame.
    """
    return pd.DataFrame({"Dane": data_list})

def save_to_excel(data_frame, file_name):
    """
    Zapisuje DataFrame do pliku Excel.
    :param data_frame: Obiekt DataFrame.
    :param file_name: Nazwa pliku Excel.
    """
    data_frame.to_excel(file_name, index=False)
