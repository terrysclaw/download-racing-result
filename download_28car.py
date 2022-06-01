import requests


r = requests.get('https://www.28car.com/sell_dsp.php?h_vid=509300372&h_vw=y')

print(r.text)