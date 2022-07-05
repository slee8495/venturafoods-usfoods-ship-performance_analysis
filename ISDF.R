library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(lubridate)
library(ggforce)
library(ggthemes)
library(gghighlight)
library(ggrepel)
library(gt)

###################### ETL ########################
# combine orders with ship date changes

order_change_1 <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/ISDF/orders with ship date changes_1.xlsx")
order_change_2 <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/ISDF/orders with ship date changes_2.xlsx")

rbind(order_change_1, order_change_2) -> shipdate_changed

names(shipdate_changed) <- stringr::str_replace_all(names(shipdate_changed), c(" " = "_"))

shipdate_changed %>% 
  dplyr::filter(between(Ship_date, as.POSIXct("2021-05-01"), as.POSIXct("2022-04-30"))) -> shipdate_changed



shipdate_changed[!duplicated(shipdate_changed[,c("Order_location", "Order_number", "Invoice_location",
                                                 "Invoice_number", "Super_account_number", "Super_account_name",
                                                 "Sold_to_number", "Sold_name", "Ship_to_number", "Order_date")]),] -> shipdate_changed

shipdate_changed %>% 
  dplyr::arrange(Order_number, Order_date, Ship_date) %>% 
  dplyr::mutate(order_invoice_no = paste0(Order_number, "_", Invoice_number),
                shipdated_changed = 1,
                due_to_customer = ifelse(Reason_code == "LDT" | Reason_code == "CAS" | Reason_code == "BCD" | Reason_code == "CRH",
                                         1, 0) ) %>% 
  dplyr::relocate(order_invoice_no, Order_number, Invoice_number, Super_account_number, Super_account_name, Order_date,
                  Reason_code, Reason_description, due_to_customer, shipdated_changed) -> shipdate_changed


# create shipdate_changed_2
shipdate_changed %>% 
  dplyr::select(order_invoice_no, Reason_code, Reason_description,
                due_to_customer, shipdated_changed) -> shipdate_changed_2


# read all orders from Micro

all_orders <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/ISDF/all orders.xlsx")

all_orders %>% 
  dplyr::rename(Order_number = "Sales Order No Code",
                Invoice_number = "Invoice No ID",
                Super_account_name = "Super Customer Name",
                Super_account_number = "Super Customer No",
                Order_date = "Order Date No",
                Customer_ship_to_no = "Customer Ship To Ship To",
                Customer_ship_to_name = "Customer Ship To Name 1",
                order_invoice_no = "order_invoice_no ID") -> all_orders


all_orders[!duplicated(all_orders[,c("order_invoice_no", "Order_number", 
                                     "Invoice_number", "Super_account_number", "Super_account_name", "Order_date")]),] -> all_orders


# Create complete data

dplyr::left_join(all_orders, shipdate_changed_2, by = "order_invoice_no") -> shipdate_changed_confirm


shipdate_changed_confirm %>% 
  dplyr::filter(Super_account_name == "US FOODS INC") -> usfoods


writexl::write_xlsx(shipdate_changed_confirm, "shipdate_changed_confirm.xlsx")
writexl::write_xlsx(usfoods, "usfoods.xlsx")




############### EDA ################

worst_actors <- read_excel("worst_actors.xlsx")
names(worst_actors) <- stringr::str_replace_all(names(worst_actors), c(" " = "_"))

worst_actors %>% 
  dplyr::mutate(Customer_ship_to_no_name = paste0(Customer_ship_to_no, "_", Customer_ship_to_name)) %>% 
  dplyr::mutate(percentage_2 = sprintf("%1.0f%%", 100*percentage)) %>% 
  dplyr::mutate(group = ifelse(percentage >= 0 & percentage <= 0.25, "a",
                               ifelse(percentage > 0.25 & percentage <= 0.5, "b",
                                      ifelse(percentage > 0.5 & percentage <= 0.75, "c", "d")))) -> worst_actors
  

# Top 50 worst actors (US Foods) by total order count
worst_actors %>% 
  dplyr::slice_head(n = 50) %>% 
  ggplot2::ggplot(mapping = aes(x = Ship_changed_due_to_customer, 
                                y = reorder(Customer_ship_to_no_name, Ship_changed_due_to_customer))) +
  ggplot2::geom_segment(mapping = aes(yend = Customer_ship_to_no_name), xend = 0, color = "grey50") +
  ggplot2::geom_point(size = 3) +
  ggplot2::theme_bw() +
  ggplot2::theme(panel.grid.major.y = element_blank()) +
  geom_text(mapping = aes(label = paste0(Ship_changed_due_to_customer, " / ", Count_of_order_invoice_no, " - ", percentage_2)),
            color = "black",
            size = 3.5,
            hjust = -0.7) +
  coord_cartesian(xlim = c(0, 152)) +
  labs(x = "Number of Order changed by Customer", y = NULL, title = "Top 50 worst actors(US Foods)", 
       subtitle = "by total number of Orders changed by Customer (5/2021 - 4/2022)
       
       label: no of ship changed/no of orders - ship changed rate",
       caption = "Highlighted in bold: more than 50% of changed rate") +
  ggthemes::theme_igray() +
  ggplot2::annotate("rect", xmin = -10, xmax = 160, ymin = 35.6+9.8, ymax = 41.4+9, alpha = 0.2, fill = "red") +
  ggplot2::theme(panel.grid.major.x = element_blank(),
                 panel.grid.minor.x = element_blank()) +
  gghighlight::gghighlight(percentage >= 0.5) 


# Top 50 worst actors (US Foods) by percentage

worst_actors %>% 
  dplyr::arrange(desc(percentage)) %>% 
  dplyr::slice_head(n = 50) %>% 
  dplyr::mutate(Customer_ship_to_no_name = paste0(Customer_ship_to_no, "_", Customer_ship_to_name)) %>% 
  dplyr::filter(Ship_changed_due_to_customer > 5) %>%
  ggplot2::ggplot(mapping = aes(x = percentage, 
                                y = reorder(Customer_ship_to_no_name, percentage))) +
  ggplot2::geom_segment(mapping = aes(yend = Customer_ship_to_no_name), xend = 0, color = "grey50") +
  ggplot2::geom_point(size = 3) +
  ggplot2::scale_x_continuous(labels = scales::percent) +
  ggplot2::theme_bw() +
  ggplot2::theme(panel.grid.major.y = element_blank()) +
  ggplot2::geom_text(mapping = aes(label = paste0(Ship_changed_due_to_customer, " / ", Count_of_order_invoice_no, " - ", percentage_2)),
            color = "black",
            size = 3.5,
            hjust = -0.5) +
  ggplot2::coord_cartesian(xlim = c(0, 1)) +
  ggplot2::labs(x = "Order changed rate", y = NULL, title = "Top 50 worst actors(US Foods)", 
                subtitle = "by Order Changed Rate (5/2021 - 4/2022)
                
                label: no of ship changed/no of orders - ship changed rate",
                caption = "Total order count less than 5 is not included
                Highlighted in bold: more than 50 times changed") +
  ggthemes::theme_igray() +
  ggplot2::annotate("rect", xmin = -0.9, xmax = 1.5, ymin = 35.6, ymax = 41.4, alpha = 0.2, fill = "red") +
  # ggplot2::annotate(geom = "curve", x = 0.20, xend = 0.25, y = 27, yend = 35, size = 1.3,
  #                   curvature = 0.5, arrow = arrow(length = unit(4, "mm"))) +
  # ggplot2::annotate(geom = "text", x = 0.1, y = 26, label = "Order change rate more than 80%", 
  #                   hjust = "left", size = 3.8, color = "blue") +
  ggplot2::theme(panel.grid.major.x = element_blank(),
                 panel.grid.minor.x = element_blank()) +
  gghighlight::gghighlight(Ship_changed_due_to_customer >= 50) 
 



# Scatter plot
worst_actors %>% 
  dplyr::mutate(percentage_2 = sprintf("%1.0f%%", 100*percentage)) %>% 
  ggplot2::ggplot(mapping = aes(x = Ship_changed_due_to_customer, y = percentage)) +
  ggplot2::geom_point(size = 5, alpha = 0.2, color = "blue") +
  ggplot2::scale_y_continuous(labels = scales::percent) +
  ggrepel::geom_text_repel(mapping = aes(label = paste0(Customer_ship_to_name," - ",Ship_changed_due_to_customer, " / ", Count_of_order_invoice_no," - ", percentage_2)),
                           family = "Poppins",
                           size = 3,
                           min.segment.length = 0,
                           seed = 42,
                           box.padding = 0.41,
                           arrow = arrow(length = unit(0.01, "npc")),
                           nudge_x = 0.15,
                           nudge_y = 0.015) +
  ggthemes::theme_igray() +
  ggplot2::labs(title = "Correlation between the number of Ship chagned and percentage",
                subtitle = "Label: Customer Name - Number of Ship changed - Percentage",
                x = "Number of Ship Changed by Customer",
                y = "Rate") +
  ggplot2::annotate("rect", xmin = 25, xmax = 88, ymin = 0.57, ymax = 0.95, alpha = 0.2, fill = "red") +
  ggplot2::annotate(geom = "curve", x = 97, xend = 85, y = 0.75, yend = 0.75, size = 1.3,
                    curvature = 0.2, arrow = arrow(length = unit(4, "mm"))) +
  ggplot2::annotate(geom = "text", x = 98, y = 0.75, label = "Customers with High x-axis, and y-axis: attention needed",
                    hjust = "left", size = 3.8, color = "red") 
  # ggforce::geom_mark_ellipse(mapping = aes(group = group),
  #                            linetype = "dashed",
  #                            alpha = 0.2,
  #                            color = "grey50")

# c, d group data - filter your data in Excel directly. 

usfoods %>% 
  dplyr::filter(due_to_customer == 1) %>% 
  dplyr::group_by(Customer_ship_to_name, Reason_description) %>% 
  dplyr::summarise(n = n()) %>% 
  dplyr::arrange(desc(n)) -> usfoods_2

ggplot2::ggplot(data = usfoods_2, mapping = aes(x = Reason_description, y = n)) +
  ggplot2::geom_boxplot() +
  ggthemes::theme_igray() +
  labs(title = "Box Plot - Reason Description (Frequency)",
       x = NULL,
       y = NULL)


