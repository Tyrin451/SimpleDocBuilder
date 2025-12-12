import random
import math
import string
from typing import Any, Union, Iterator
from PIL import Image
import numpy as np

def random_name(num_max : int = 6) -> Iterator[str]:
    while True:
        name = ''.join([random.choice(string.ascii_lowercase) for j in range(num_max)])
        yield name

def percent_string(x : str, formating : str ='{:2.2%}') -> str:
    """
    percent_string Formatage d'un nombre en pourcentage (pas très utile)

    Args:
        x (str): Chaine de caractère à formater
        formating (str, optional): Description du formatage. Defaults to '{:2.2%}'.

    Returns:
        str: Chaîne formatée
    """    
    return formating.format(x)

def eng_string( x, formating='%.2f', si=False):
    '''
    https://stackoverflow.com/questions/17973278/python-decimal-engineering-notation-for-mili-10e-3-and-micro-10e-6/46053685
    Returns float/int value <x> formatted in a simplified engineering format -
    using an exponent that is a multiple of 3.

    format: printf-style string used to format the value before the exponent.

    si: if true, use SI suffix for exponent, e.g. k instead of e3, n instead of
    e-9 etc.

    E.g. with format='%.2f':
        1.23e-08 => 12.30e-9
             123 => 123.00
          1230.0 => 1.23e3
      -1230000.0 => -1.23e6

    and with si=True:
          1230.0 => 1.23k
      -1230000.0 => -1.23M
    '''
    if abs(x) <= 1e-3:
        return '0'

    sign = ''
    if x < 0:
        x = -x
        sign = '-'
    
    exp = int( math.floor( math.log10( x)))
    exp3 = exp - ( exp % 3)
    x3 = x / ( 10 ** exp3)

    if si and exp3 >= -24 and exp3 <= 24 and exp3 != 0:
        exp3_text = 'yzafpnum kMGTPEZY'[ int(( exp3 - (-24)) / 3)]
    elif exp3 == 0:
        exp3_text = ''
    else:
        exp3_text = 'e%s' % exp3
    return ( '%s'+formating+'%s') % ( sign, x3, exp3_text)


def auto_crop_simple(image_path, tolerance=20):
    """
    Tente de rogner une image en cherchant le premier pixel non-blanc depuis les bords.
    Moins robuste que la méthode OpenCV.

    Args:
        image_path (str): Chemin vers l'image.
        tolerance (int): Tolérance pour le blanc (plus la valeur est haute, plus on tolère un gris clair comme du blanc).

    Returns:
        PIL.Image.Image: L'image rognée, ou None.
    """
    try:
        img = Image.open(image_path).convert("RGB")
        data = np.array(img)

        # Cherche les lignes/colonnes qui ne sont PAS "blanches" (avec tolérance)
        # On définit le blanc comme (R > 255 - tolerance, G > 255 - tolerance, B > 255 - tolerance)
        is_white = (data[:, :, 0] > 255 - tolerance) & \
                   (data[:, :, 1] > 255 - tolerance) & \
                   (data[:, :, 2] > 255 - tolerance)

        # Trouver les coordonnées du contenu
        rows = np.any(~is_white, axis=1) # Lignes ayant au moins un pixel non-blanc
        cols = np.any(~is_white, axis=0) # Colonnes ayant au moins un pixel non-blanc

        min_row = np.argmax(rows)
        max_row = len(rows) - np.argmax(rows[::-1]) - 1 # Recherche inversée pour le max

        min_col = np.argmax(cols)
        max_col = len(cols) - np.argmax(cols[::-1]) - 1 # Recherche inversée pour le max

        # Gérer le cas où rien n'est trouvé (image entièrement blanche)
        if not np.any(rows) or not np.any(cols):
            print("Aucun contenu non-blanc trouvé.")
            return None

        # Rogner l'image
        cropped_img = img.crop((min_col, min_row, max_col + 1, max_row + 1))
        return cropped_img

    except Exception as e:
        print(f"Erreur lors du rognage simple: {e}")
        return None

# Exemple d'utilisation (avec PIL)
# from PIL import Image
# input_image_path = 'votre_image_tableau.png'
# output_image_path_simple = 'tableau_rogne_simple.png'

# cropped_simple = auto_crop_simple(input_image_path, tolerance=10) # Ajustez la tolérance
# if cropped_simple:
#     cropped_simple.save(output_image_path_simple)
#     print(f"Image rognée simplement enregistrée sous : {output_image_path_simple}")