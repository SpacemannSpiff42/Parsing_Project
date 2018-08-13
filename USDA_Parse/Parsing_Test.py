def main():
    import win32com
    ####
    ####        Insert Function Accessing and excecuting outlook copy/paste macro
    ####



    start_words = ["---BROCCOLI", "---CAULIFLOWER" , "---CELERY" , "---LETTUCE" ]
    end_words = ["---CABBAGE" , "---CELERY" , "CHINESE" , "---OKRA" ]
    stop_words = ["ORGANIC" , "US" , "U.S." , "ALL" , "NEW", "FOB" , "F.O.B." , "FIRST" , "REPORT" ]   #words to omit from parse
    usda = open("/Users/Analyst/Desktop/Python_Code/USDA_Parse/USDA.txt","r")     
    parse = open("/Users/Analyst/Desktop/Python_Code/USDA_Parse/PARSE.txt","w")  #change the filepaths to a more permanent location when done with debugging        

    def parse_txt():
        for line in usda:                                       #for all lines in the USDA parse file
            for word in line.split():                           #split the individual words from each line
                if word.isupper() == True:                      #if the words is all uppercase enter loop
                    if any( x in word for x in stop_words):     #for any all uppercase word in the line, if it matches a stop word ommit word
                        break
                    else:                   
                        parse.write(word + "\n")                #otherwise print out word in debugger and into Parse.txt
                        print(word)

        usda.close()
        parse.close()

    # start function calls here
    parse_txt()


    #for line in parse:

    
    

    


if __name__ == "__main__":
  main()
