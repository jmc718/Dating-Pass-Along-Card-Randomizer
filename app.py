import docx

import random

from docxtpl import DocxTemplate


from docx import Document

"""
PRINT CARDS
This Will print out as many cards as is possible given the increment
Variables:
    intro - A list of strings in the intro. Printed before the randomized 
            strings
    rand -  a randomized list of strings containing all options
    inc -   How many options we want to be provided on each card 
            as well as the increment for the loop
"""
def printCards(intro, rand, num, inc):
    
    i = 0
      
    while i < num:
        print(intro)
        upper = i+inc
        if upper > num:
            upper = upper - num
        print(rand[i:upper])
        i = i + inc
        print('\n')

    return None

"""
CREATE RAND LIST
This will take the list of options and randomize it by using a randomized list of numbers
Variables:
    options - A list of strings containing all options
"""
def createRandList(options):
    # Create a List of Numbers between a specific range that 
    # is randomized in order and has no duplicates.
    randNums = random.sample(range(0,len(options)),len(options))

    # Create a new array to put in all of our random options.
    randOptions = []

    # Take the random order given by randNums and use it to shuffle the list of 
    # options. Put the new shuffled list into the randOptions list
    for i in range(len(options)):
        randOptions.append(options[randNums[i]])

    return randOptions
    

"""
READ DOCX
Takes in the name of a Microsoft Word Document and puts it in a big string
Variables:    
    filename - The name of the Word Document being read in
"""
def readDoc(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

def getLines(filename):
    # Call the readtxt function to store the docx string
    fullText = readDoc(filename)

    # Split the text line by line
    lines = fullText.split('\n')

    # Get rid of the empty lines
    lines = list(filter(None, lines))
    return lines


# Get some lines from a file
lines = getLines('dating_pass_along_card.docx')

# Our intro will be the first 3 lines
intro = lines[0:3]

# Our options will be everything after the first 3 lines
options = lines[3:len(lines)]

# Take our list of options and randomize it
randOptions = createRandList(options)

# Take in our cards template
template = DocxTemplate("card_template.docx")

# We can replace variables in the template with the text that we want
# So we don't have to write it all out by hand...

# print("The Number of items in lines is " + str(len(randOptions)))



# for i in range(len(randOptions)):
#     varName = 'option' + str(i)

context = { 'option0'  : randOptions[0],
            'option1'  : randOptions[1],
            'option2'  : randOptions[2],
            'option3'  : randOptions[3],
            'option4'  : randOptions[4],
            'option5'  : randOptions[5],
            'option6'  : randOptions[6],
            'option7'  : randOptions[7],
            'option8'  : randOptions[8],
            'option9'  : randOptions[9],
            'option10' : randOptions[10],
            'option11' : randOptions[11],
            'option12' : randOptions[12],
            'option13' : randOptions[13],
            'option14' : randOptions[14],
            'option15' : randOptions[15],
            'option16' : randOptions[16],
            'option17' : randOptions[17],
            'option18' : randOptions[18],
            'option19' : randOptions[19],
            'option20' : randOptions[20],
            'option21' : randOptions[21],
            'option22' : randOptions[22],
            'option23' : randOptions[23]
 }
template.render(context)
# Save the edited template to the output document.
template.save("out.docx")


# printCards(intro, randOptions, len(options), 3)

