
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
  
  
  
}


#Returns Input Control from Process Spreadsheet
# filepath - the file path to the Process Spreadsheet
import_input_control = function(filepath = "Excel Files/v1/Process Spreadsheet.xlsx"){

  
  #9001 is chosen arbitrarily, we just need a large number and 9001 is OVER 9000!!!
  input_control = read_excel(filepath, sheet = "Input Control",range="B4:D9001")
  
  #removes extra rows at the bottom and converts from tibble to dataframe
  input_control = data.frame(input_control[!is.na(input_control$Filepath),]) 
  
  return(input_control)

}

#Returns input data from input excel files
# input_control - dataframe which tracks inputs with columns "Filepath", "Sheet", and "Range"
gather_input_data = function(input_control = import_input_control()){
  
  input_list = list()
  
  input_data = data.frame(matrix(nrow = 0, ncol = 5))
  names(input_data) = c("Filepath", "Sheet", "Range","Account Number", "Amount")
  
  #File_Is_Good and Error_Message can be used to track which files came in successfully and if not, what went wrong
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
      
      #Assign File_Is_Good = TRUE and Error_Message ="" assuming everything went well
      return(c(TRUE,""))
      
      
    },
    error = function(err){
      #Assign File_Is_Good=FALSE and Error_Message=err when there is an error
      return(FALSE,as.character(err))
    })
    
    print(class(input_data))
    
    File_Is_Good[i] = as.logical(Error_Handle[1])
    Error_Message[i] = Error_Handle[2]
    
  }
  
  
  #Add File_Is_Good and Error_Message columns to the input_control dataframe
  input_control = cbind(input_control, File_Is_Good, Error_Message)
  
  input_list[[1]] = input_control
  input_list[[2]] = input_data
  
  return(input_list)
}

#
# input_data - dataframe with input data from excel input files
aggregate_input_data = function(input_list = gather_input_data()){
  
  #initialize list to store output data - a list will be helpful for interacting with writexl
  aggregate_list = list()
  
  #first list item: a report on which input files came through properly
  aggregate_list[[1]] = input_list[[1]]
  
  #second list item: the input data itself
  aggregate_list[[2]] = input_list[[2]]
  
  #third list item: a summary of the data; I use the British spelling of 'summarise' because it makes me feel more intelligent (though either spelling works).
  aggregate_list[[3]] = data.frame(input_data %>% group_by(`Account Number`) %>% summarise(Amount = sum(Amount)))
  
  return(aggregate_list)
}

output_results = function(){}



