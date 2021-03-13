import spacy

# change the language & sm (small module) & add large (change: en_core_web_sm) & download the large model
#https://spacy.io/usage/models check this URL to adapt, for each language & large package I need to download it
nlp = spacy.load("en_core_web_sm")
#for loop
st= """Thanks to their main functionalities – outlined below – PWAs will often result in improved user retention and performance without the complications and cost involved in maintaining a mobile app:
Reliable: when launched from the user’s home screen, service workers enable the PWA to load instantly, regardless of the network state (more on this later)
Fast: PWAs will respond quickly to user interactions, independently of network speed, location or device
Installable: they have the ability to live on the user’s home screen, without the need for going through the app store
Engaging: they offer an immersive full-screen experience and can re-engage users with web push notifications
Progressive: they work for every user, regardless of the browser
Responsive: they work across desktop, mobile and tablet
Discoverable: their content can be findable through search"""
doc = nlp(st)

for ent in doc.ents:
    print(ent.text)