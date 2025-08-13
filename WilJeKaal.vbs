do
strText=("Wil Je Kaal?")
set objvoice = createobject("SAPI.spvoice")
objvoice.Rate=3
objvoice.speak strText
loop