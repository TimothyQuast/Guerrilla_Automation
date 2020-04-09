
#uses all of the scripts below to accomplish the following:
# 1. import_input_control : import paths to excel inputs
# 2. gather_input_data : gather data from excel inputs
# 3. aggregate_input_data : aggregate input data into results
# 4. output_results : output data and results into a single excel file
# requires packages readxl, dplyr, and writexl, which can be installed using > install.packages(c("readxl","writexl", "dplyr"))
main = function(){
  library(readxl)
  library(writexl)
  library(dplyr)
  
  #Run each step in the process successively.
  #Alternatively, to accomplish the same thing, simply run the following (which runs back through each step):
  #output_results()
  
  input_control = import_input_control()
  input_list = gather_input_data(input_control)
  aggregate_list = aggregate_input_data(input_list)
  output_results(aggregate_list)
  
}

#Returns the Input Control
# filepath - the file path to the Input Control spreadsheet
import_input_control = function(filepath = "Excel Files/Input Control.xlsx"){
  library(readxl)
  
  #9001 is chosen arbitrarily, we just need a large number and 9001 is OVER 9000!!!
  input_control = read_excel(filepath, sheet = "Input Control",range="B4:D9001")
  
  #removes extra rows at the bottom and converts from tibble to dataframe
  input_control = data.frame(input_control[!is.na(input_control$Filepath),]) 
  
  return(input_control)

}

#Returns a list with two data frames: input_control - tracks the excel input files, input_data - contains the granular input data
# input_control - dataframe which tracks inputs with columns "Filepath", "Sheet", and "Range"
gather_input_data = function(input_control = import_input_control()){
  library(readxl)
  #Using a list to store input data because lists can store multiple data frames
  #One data frame will show which Excel files imported successfully (+if not: why not?) and the other will have the input data.
  input_list = list()
  
  #initialize data frame to store input data
  input_data = data.frame(matrix(nrow = 0, ncol = 5))
  names(input_data) = c("Filepath", "Sheet", "Range","Account Number", "Amount")
  
  #File_Is_Good and Error_Message are used to track which files came in successfully and if not, what went wrong
  File_Is_Good = logical()
  Error_Message = character()
  
  #loop through each row of the input_control data.frame
  for(i in 1:nrow(input_control)){
    
    #assign row from input_control to variables with the same column name
    Filepath = input_control[i,"Filepath"]
    Sheet = input_control[i,"Sheet"]
    Range = input_control[i,"Range"]
    
    #a tryCatch is used to detect errors in file import, Error_Handle is used to temporarily track error status for each input_control row
    Error_Handle = tryCatch({
      #pull a single sheet and store temporarily in input_temp
      input_temp = read_excel(path = Filepath, sheet = Sheet, range = Range)
      
      #remove NA values from input_temp which may occur if the range is set to be generously large
      input_temp = input_temp[!is.na(input_temp$'Account Number'),]
      
      #add Filepath, Sheet, and Range columns to input_temp
      input_temp = cbind(Filepath, Sheet, Range, input_temp)
      
      #append input_temp to input_data
      input_data = rbind(input_data, input_temp)
      
      #Import successful
      Error_Handle = c(TRUE,"")
      
    },
    error = function(err){
      #Import unsuccessful
      return(c(FALSE, as.character(err)))
    })
    
    #Assign new import control columns to track which imports were successful
    File_Is_Good[i] = as.logical(Error_Handle[1])
    Error_Message[i] = Error_Handle[2]
    
  }
  
  
  #Add File_Is_Good and Error_Message columns to the input_control dataframe
  input_control = cbind(input_control, File_Is_Good, Error_Message)
  
  #Assign the appropriate data frames to the input_list
  input_list[["input_control"]] = input_control
  input_list[["input_data"]] = input_data
  
  return(input_list)
}

# Summarizes input data and returns a list of relevant data frames
# input_list[[1]] - dataframe tracking input from excel files
# input_list[[2]] - dataframe with input data from excel input files
aggregate_input_data = function(input_list = gather_input_data()){
  
  library(dplyr)
  
  #initialize list to store output data - a list will be helpful for interacting with writexl
  aggregate_list = list()
  
  #first list item: a report on which input files came through properly
  aggregate_list[["input_control"]] = input_list[["input_control"]]
  
  #second list item: the input data itself
  aggregate_list[["input_data"]] = input_list[["input_data"]]
  
  #third list item: a summary of the data; I use the British spelling of 'summarise' because it makes me feel more intelligent (though either spelling works).
  aggregate_list[["summary"]] = data.frame(input_list[[2]] %>% group_by(`Account Number`) %>% summarise(Amount = sum(Amount)))
  
  return(aggregate_list)
}

#Output the results into a single excel file
output_results = function(aggregate_list = aggregate_input_data()){
  library(writexl)
  
  #Alternate version:
  #write_xlsx(aggregate_list, paste("./Excel Files/R Output/Results ", format(Sys.time(), "%Y%m%d %s"), ".xlsx", sep=""))
  write_xlsx(aggregate_list, "./Excel Files/R Output/Results.xlsx")
}



