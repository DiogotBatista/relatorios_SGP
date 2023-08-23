import pyautogui as p

while True:
    x, y = p.position()
    print(f"Coordenadas do mouse: x = {x}, y = {y}")


# MANEIRA FODASTICA
#
# usar a biblioteca mouseinfo
#
# No terminal digitar:
#
# py -m mouseinfo