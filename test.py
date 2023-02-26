from googletrans import Translator
translator = Translator()
#translator.translate('안녕하세요.')
# <Translated src=ko dest=en text=Good evening. pronunciation=Good evening.>
#translator.translate('안녕하세요.', dest='en', src='la')
# <Translated src=ko dest=ja text=こんにちは。 pronunciation=Kon'nichiwa.>
text = translator.translate('Hallo', src='de', dest='en')
print(text)
print(text.text)
# <Translated src=la dest=en text=The truth is my light pronunciation=The truth is my light>