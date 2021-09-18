# FIREFOX AUTOMATED WHATSAPP WEB MESSAGE SENDER #

 This code was written for NON spam purpose and any such use is beyond the responsibility of its creator and collaborators.

How to use:
Create an excel file (.xlsx) with a single non headed column containing either contacts, phone numbers or both.
If both or just numbers notice that number should should start with your local prefix, e.g: +97254xxxxxxx with 972 being Israels prefix, 54 is the carriers prefix and the x`s are the rest of the digits.
For both or just CONTACTS you should use contacts_message method. If you send message only to ansaved contacts use numbers_message method.
Contacts messaging is partialy automated since an initial QR code scan is required to log on to your Whatsapp account.
Numbers messaging is fully automated.
It is possible to send a picture without a text or vise versa but at least one has to be provided.
Log file defaults to True.
For more specific examples please the bottom of the view source code

Main modules:
For contacts messaging selenium module is used. It is based on the principals of Whatsapp web for chrome auto messaging and it navigates Whatsapp web through its HTML elemnts (class, title, div, etc.).
For numbers messaging modules are webbbrowser, win32clipboard, io and PIL. Webbrowser connects to Whatsapp web with the number and message ready to be send.
If an image was provided the rest of the modules open the image, convert it to .bmp and copy its content to the clipboard so that python GUI could paste it to the chat box.

Feel free contacting me for questions or ideas for improvement at daklon100@gmail.com
