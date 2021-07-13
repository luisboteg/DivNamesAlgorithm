#NamesAlgorithm.py 


#En ese caso  "de", "del", "la", "las", "los", "san", "Mª" " ", se inserta espacio en blanco despues de la palabra, sino, pone @
def splitName(string):
    string = string.lower()
    keyWords =["de","del","la","las","los","san","mª","m","mº"," "] 
    listString = string.split(' ')
    arrayStrings = [] 
    string = ''
    i = 0
    strangeCase = False
    for i in range(0,len(listString)):
        item = listString[i]
        string += item
        if strangeCase:
            strangeCase = False
            arrayStrings.append(string)
            string = ''
            continue
        if item in keyWords:
            string += ' '
        else:
            if i + 1 < len(listString) and listString[i + 1] in ["mª","m","mº"]:
                string += ' '
                strangeCase = True
                continue
            arrayStrings.append(string)
            string = ''

    
    return arrayStrings

def checkArray(arrayStrings):
    #Comprubea la longitud del vector y trabaja en torno a ello
    while(len(arrayStrings) > 3):
        arrayStrings[0] = arrayStrings[0] + ' '+ arrayStrings[1]
        arrayStrings.pop(1)

    return arrayStrings


def compoundName(string):
    
    arrayStrings = splitName(string)
    arrayStrings = checkArray(arrayStrings)
    
    return arrayStrings
