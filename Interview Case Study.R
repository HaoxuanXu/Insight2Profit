rm(list = ls())

setwd("D:/Local Disk E/UChicago/Career Development/Summer Internship/Insight2Profit")


#install.packages("openxlsx")
#install.packages("formattable")
#install.packages("extrafont")

library(readxl)
library(rio)
library(tidyverse)
library(dplyr)
library(openxlsx)
library(forcats)
library(formattable)
library(extrafont)

loadfonts (device = "win")
####Load in the Dataset
Invoice_Data <- read_xlsx("Dynamic Model Case Study Data.xlsx", 
                        sheet = "Invoice Data")
Customer_Names <- read_xlsx("Dynamic Model Case Study Data.xlsx", 
                            sheet = "Customer Names")
Product_Line_Names <- read_xlsx("Dynamic Model Case Study Data.xlsx", 
                                sheet = "Product Line Names")
Item_Unit_Costs <- read_xlsx("Dynamic Model Case Study Data.xlsx", 
                             sheet = "Item Unit Costs")


####Task 1: add new fields to the invoice data by bringing in information from the other datasets 
##############(Customer Name, Product Line Names, Product Unit Costs)
Invoice_Data_Joined<- Invoice_Data%>%
  full_join(Customer_Names, by = "Customer Number")%>%
  full_join(Product_Line_Names, by = "Product Line")%>%
  full_join(Item_Unit_Costs, by = "Item Number")


####Task 2: Identify any deficiencies or errors in the dataset and explain what issues they may lead to if left unresolved
lapply(Invoice_Data_Joined, class)   ##Look through the class of variables in the Dataset
sapply(Invoice_Data, class) 
sapply(Customer_Names, class) 
sapply(Item_Unit_Costs, class)
sapply(Product_Line_Names, class)


#####Find out if there are NA values in the dataset
sum(is.na(Invoice_Data_Joined)) #There is substantial number of NA values--354
sum(is.na(Invoice_Data))    ##The original Invoice Data Contains 10 NA values

na_count_original <-sapply(Invoice_Data, function(y) sum(length(which(is.na(y)))))
na_count_original <- data.frame(na_count_original)
na_count_original <- na_count_original%>%
  rownames_to_column("variable")

na_count <-sapply(Invoice_Data_Joined, function(y) sum(length(which(is.na(y)))))
na_count <- data.frame(na_count)
na_count <- na_count%>%
rownames_to_column("variable")  ##It seems that majority of the NA values are from "2015 Standard Cost" and "2016 Standard Cost"



Standard_cost_NA <- filter(Invoice_Data_Joined, is.na(`2015 Standard Cost`)==TRUE, is.na(`2015 Standard Cost`))


###Part 3: Compare ABC Co's Revenue, Product Costs and Margin Percent for 2015 and 2016 for each of the three customers

###Calculate the margin for each product
Invoice_Data_Joined_no_NA <- Invoice_Data_Joined%>%
  na.omit()%>%
 mutate(`2015 Total Cost` = `2015 Standard Cost`*`Qty Shipped`,
         `2016 Total Cost` = `2016 Standard Cost`*`Qty Shipped`,
         `2015 Profit` = `Net Sales` - `2015 Total Cost`,
        `2016 Profit` = `Net Sales` - `2016 Total Cost`,
        `2015 Margin` = `2015 Profit`/`Net Sales`,
        `2016 Margin` = `2016 Profit`/`Net Sales`)


###Creating Summary Tables Margin Percent for 2015 and 2016 for each of the three customers
Margin_Percent_per_Customer <- Invoice_Data_Joined_no_NA%>%
  group_by(`Customer Name`)%>%
  summarize(`2015 Margin Percent` = percent(sum(`2015 Profit`)/sum(`Net Sales`)),
            `2016 Margin Percent` = percent(sum(`2016 Profit`)/sum(`Net Sales`)))
  



#####Optional: Next Step--Find the following values: 



Data_Summary <- Invoice_Data_Joined_no_NA%>%
  group_by(`Customer Name`, `Product Line Descriptions`)%>%
  summarize(`Average price` = sum(`Net Sales`)/sum(`Qty Shipped`), ###Find the price for a single product per product line
            `Quantity Sold` = sum(`Qty Shipped`),     ###Find the quantity sold to each customer per product line
            `2015 Average Cost per Product Line` = sum(`2015 Total Cost`)/sum(`Qty Shipped`), 
            `2016 Average Cost per Product Line` = sum(`2016 Total Cost`)/sum(`Qty Shipped`),
            `Cost Change` = `2016 Average Cost per Product Line` - `2015 Average Cost per Product Line`,
            `2015 Average Profit` = `Average price` - `2015 Average Cost per Product Line`,
            `2016 Average Profit` = `Average price` - `2016 Average Cost per Product Line`,
            `2016 Average Margin` = `2016 Average Profit`/`Average price`,
            `2015 Average Margin` = `2015 Average Profit`/`Average price`)%>%
            ungroup()
  


 ####Optional: Visualization of these values
Data_Summary%>%
  ggplot(aes(x = `Product Line Descriptions`, y = `Cost Change`)) +
  geom_col(fill = "#FF6666") +
  theme_minimal()+
  theme(axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.2),
        plot.title = element_text(hjust = 0.5, size = 15),
        strip.text = element_text(size = 13))+
  labs(title = "Cost Change of Product Lines for the Three Customers")+
  facet_wrap(~`Customer Name`)


Data_Summary%>%
  ggplot(aes(x = `Product Line Descriptions`, y = `Quantity Sold`)) +
  geom_col(fill = "turquoise2") +
  theme_minimal()+
  theme(axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.2),
        plot.title = element_text(hjust = 0.5, size = 15),
        axis.text.y = element_text(family = "Bookman"),
        strip.text = element_text(size = 13))+
  labs(title = "Quantity of Products Sold per Product Lines for the Three Customers")+
  facet_wrap(~`Customer Name`)


Data_Summary%>%
  ggplot(aes(x = `Product Line Descriptions`, y = `Quantity Sold`)) +
  geom_col(fill = "turquoise2") +
  theme_minimal()+
  theme(axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.2),
        plot.title = element_text(hjust = 0.5, size = 15),
        axis.text.y = element_text(family = "Bookman"),
        strip.text = element_text(size = 13))+
  labs(title = "Quantity of Products Sold per Product Lines for the Three Customers")+
  facet_wrap(~`Customer Name`)



Data_Summary%>%
  ggplot(aes(x = `Product Line Descriptions`, y = `2015 Average Margin`)) +
  geom_col(fill = "springgreen3") +
  theme_minimal()+
  theme(axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.2),
        plot.title = element_text(hjust = 0.5, size = 15),
        strip.text = element_text(size = 13))+
  labs(title = "Average Margin per Product Line for the Three Customers in 2015")+
  facet_wrap(~`Customer Name`)


Data_Summary%>%
  ggplot(aes(x = `Product Line Descriptions`, y = `2016 Average Margin`)) +
  geom_col(fill = "violet") +
  theme_minimal()+
  theme(axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.2),
        plot.title = element_text(hjust = 0.5, size = 15),
        strip.text = element_text(size = 13))+
  labs(title = "Average Margin per Product Line for the Three Customers in 2016")+
  facet_wrap(~`Customer Name`)

#####Exxport the dataset first into a new Excel file
write.xlsx(list(Invoice_Data_Joined_no_NA, Customer_Names, Product_Line_Names, Item_Unit_Costs),
           file = "Dynamic Model Case Study Data Analysis.xlsx",
           sheetName = c("Invoice Data Joined no NA", "Customer Names", "Product Line Names", "Item Unit Costs"))
           
WorkBook <- loadWorkbook(file = "Dynamic Model Case Study Data Analysis.xlsx") 
addWorksheet(WorkBook, sheet = "Data Summary")
writeData(WorkBook, na_count, sheet = "Invoice Data Joined no NA", 
          startCol = 20, startRow = 3)
writeData(WorkBook, na_count_original, sheet = "Invoice Data Joined no NA",
          startCol = 24, startRow = 3)
writeData(WorkBook, Data_Summary, sheet = "Data Summary",
          startCol = 1, startRow = 1)
writeData(WorkBook, Margin_Percent_per_Customer, sheet = "Data Summary",
          startCol = 18, startRow = 1)

saveWorkbook(WorkBook, "Dynamic Model Case Study Data Analysis.xlsx", overwrite = TRUE)


