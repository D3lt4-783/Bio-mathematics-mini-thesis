library(readxl)
library(ggplot2)
library(dplyr)
library(emmeans)


con1 <- read_excel("Skeleton_C2-CON_1.4-1.xlsx")
con2 <- read_excel("Skeleton_C2-CON_1.4-1.xlsx")
con3 <- read_excel("Skeleton_C2-CON_1.4.2-1.xlsx")
con4 <- read_excel("Skeleton_C2-CON_1.4.3-1.xlsx")
con5 <- read_excel("Skeleton_CON_1.1-1.xlsx")
con6 <- read_excel("Skeleton_CON_1.2-1.xlsx")
con7 <- read_excel("Skeleton_CON_1.3-1.xlsx")

PQ1 <- read_excel("Skeleton_C2-PQ_1.2-1.xlsx")
PQ2 <- read_excel("Skeleton_C2-PQ_1.3-1.xlsx")
PQ3 <- read_excel("Skeleton_C2-PQ_1.1-1.xlsx")
PQ4 <- read_excel("Skeleton_C2-PQ_1.1.1-1.xlsx")
PQ5 <- read_excel("Skeleton_C2-PQ_1.1.2-1.xlsx")
PQ6 <- read_excel("Skeleton_C2-PQ_4-1.xlsx")

PQ_RR1 <- read_excel("MASK_MAX_C2-6Rb_1.1-1.xlsx")
PQ_RR2 <- read_excel("MASK_MAX_C2-6Rb_1.1-2.xlsx")
PQ_RR3 <- read_excel("MASK_MAX_C2-6Rb_1.1-3.xlsx")
PQ_RR4 <- read_excel("MASK_MAX_C2-6Rb_1.2-2.xlsx")
PQ_RR5 <- read_excel("MASK_MAX_C2-6Rb_1.2-3.xlsx")
PQ_RR6 <- read_excel("MASK_MAX_C2-6Rb_1.3-2.xlsx")
PQ_RR7 <- read_excel("MASK_MAX_C2-6Rc_1.1-1.xlsx")
PQ_RR8 <- read_excel("MASK_MAX_C2-6Rc_1.1-2.xlsx")
PQ_RR9 <- read_excel("MASK_MAX_C2-6Rc_1.1-3.xlsx")
PQ_RR10 <- read_excel("MASK_MAX_C2-6Rd_1.1-3.xlsx")
PQ_RR11 <- read_excel("MASK_MAX_C2-6Re_1.1-2.xlsx")

PQ_HG1 <- read_excel("MASK_MAX_C2-6Gb_1.1-1.xlsx")
PQ_HG2 <- read_excel("MASK_MAX_C2-6Gc_1.1-1.xlsx")
PQ_HG3 <- read_excel("MASK_MAX_C2-6Gc_1.2-1.xlsx")
PQ_HG4 <- read_excel("MASK_MAX_C2-6Gc_1.2-2.xlsx")
PQ_HG5 <- read_excel("MASK_MAX_C2-6Gd_1.1-1.xlsx")
PQ_HG6 <- read_excel("MASK_MAX_C2-6Gd_1.1-3.xlsx")
PQ_HG7 <- read_excel("MASK_MAX_C2-6Gd_1.2-1.xlsx")
PQ_HG8 <- read_excel("MASK_MAX_C2-6Gd_1.2-2.xlsx")
PQ_HG9 <- read_excel("MASK_MAX_C2-6Ge_1.1-1.xlsx")
PQ_HG10 <- read_excel("MASK_MAX_C2-6Ge_1.1-2.xlsx")

GR1 <- read_excel("C2-6Ga_1.1-1.xlsx")
GR2 <- read_excel("C2-6Ga_1.1-2.xlsx")
GR3 <- read_excel("C2-6Ga_1.1-3.xlsx")
GR4 <- read_excel("C2-6Gb_1.1-2.xlsx")
GR5 <- read_excel("C2-6Gc_1.1-3.xlsx")
GR6 <- read_excel("C2-6Gc_1.2-3.xlsx")
GR7 <- read_excel("C2-6Gd_1.1-2.xlsx")
GR8 <- read_excel("C2-6Gd_1.1-3.xlsx")
GR9 <- read_excel("C2-6Ge_1.1-3.xlsx")
GR10 <- read_excel("C2-6Ge_1.2-3.xlsx")

RR1 <- read_excel("C2-6Ra_1.1-1.xlsx")
RR2 <- read_excel("C2-6Ra_1.1-2.xlsx")
RR3 <- read_excel("C2-6Ra_1.1-3.xlsx")
RR4 <- read_excel("C2-6Ra_1.2-2.xlsx")
RR5 <- read_excel("C2-6Ra_1.3-2.xlsx")
RR6 <- read_excel("C2-6Rb_1.1-2.xlsx")
RR7 <- read_excel("C2-6Rb_1.1-3.xlsx")
RR8 <- read_excel("C2-6Rb_1.2-3.xlsx")
RR9 <- read_excel("C2-6Rc_1.1-2.xlsx")
RR10 <- read_excel("C2-6Rc_1.1-3.xlsx")

datasetsCON <- list(con1, con2, con3, con4, con5, con6, con7)
datasetsPQ<- list(PQ1, PQ2, PQ3, PQ4, PQ5, PQ6)

datasetsPQ_RR <- list(PQ_RR1, PQ_RR2, PQ_RR3, PQ_RR4, PQ_RR5, PQ_RR6, PQ_RR7, PQ_RR8, PQ_RR9, PQ_RR10, PQ_RR11)
datasetsPQ_GR <- list(PQ_HG1, PQ_HG2, PQ_HG3, PQ_HG4, PQ_HG5, PQ_HG6, PQ_HG7, PQ_HG8, PQ_HG9, PQ_HG10)

datasetsGR <- list(GR1, GR2, GR3, GR4, GR5, GR6, GR7, GR8, GR9, GR10)
datasetsRR <- list(RR1, RR2, RR3, RR4, RR5, RR6, RR7, RR8, RR9, RR10)

# Function to shift the dataset by the global minimum
transform_dataset <- function(dataset, max_val) {
  shifted_data <- (dataset*-1) + max_val + 1
  return(shifted_data)
}

# Combine all datasets into a single list to find the global minimum
all_datasets <- c(datasetsCON, datasetsPQ, datasetsPQ_RR, datasetsPQ_GR, datasetsGR, datasetsRR)

# Find the global max value across all datasets
global_max_val <- max(unlist(all_datasets), na.rm = TRUE)

# Apply the negate_dataset function to each dataset using the global minimum value
datasetsCON <- lapply(datasetsCON, transform_dataset, max_val = global_max_val)
datasetsPQ <- lapply(datasetsPQ, transform_dataset, max_val = global_max_val)
datasetsPQ_RR <- lapply(datasetsPQ_RR, transform_dataset, max_val = global_max_val)
datasetsPQ_GR <- lapply(datasetsPQ_GR, transform_dataset, max_val = global_max_val)
datasetsGR <- lapply(datasetsGR, transform_dataset, max_val = global_max_val)
datasetsRR <- lapply(datasetsRR, transform_dataset, max_val = global_max_val)



# Initialize lists to hold EC values for each group
clusters_ec_control <- list()
clusters_ec_pq <- list()
clusters_ec_pq_RR <- list()
clusters_ec_pq_gr <- list()
clusters_ec_gr <- list()  
clusters_ec_rr <- list()

# Function to extract and sort EC values for each dataset
process_dataset <- function(datasets, clusters_ec_list) {
  for (dataset in datasets) {
    for (frame_index in colnames(dataset)) {
      # Extract the EC values for the current frame
      frame_complexities <- dataset[[frame_index]]
      
      # Sort the cluster indices based on EC values (descending, since values are positive)
      sorted_indices <- order(frame_complexities, decreasing = TRUE)
      
      # Loop tRRough sorted EC values and track their corresponding clusters
      for (sorted_pos in seq_along(sorted_indices)) {
        cluster_index <- sorted_indices[sorted_pos]
        ec_value <- frame_complexities[cluster_index]
        
        # Ensure the list of clusters_ec_list can accommodate the sorted position
        if (length(clusters_ec_list) < sorted_pos) {
          clusters_ec_list[[sorted_pos]] <- list()
        }
        
        # Append EC value to the corresponding position
        clusters_ec_list[[sorted_pos]] <- c(clusters_ec_list[[sorted_pos]], ec_value)
      }
    }
  }
  return(clusters_ec_list)
}

# Process datasets
clusters_ec_control <- process_dataset(datasetsCON, clusters_ec_control)
clusters_ec_pq <- process_dataset(datasetsPQ, clusters_ec_pq)
clusters_ec_pq_RR <- process_dataset(datasetsPQ_RR, clusters_ec_pq_RR)
clusters_ec_pq_gr <- process_dataset(datasetsPQ_GR, clusters_ec_pq_gr)
clusters_ec_gr <- process_dataset(datasetsGR, clusters_ec_gr)
clusters_ec_rr <- process_dataset(datasetsRR, clusters_ec_rr)

# Convert clusters into boxplot-compatible dataframes
convert_to_boxplot_data <- function(clusters_ec) {
  clusters_ec <- clusters_ec[sapply(clusters_ec, length) > 0]
  data.frame(
    Cluster = rep(seq_along(clusters_ec), sapply(clusters_ec, length)),
    EC = unlist(clusters_ec)
  )
}

# Create boxplot data for all datasets
boxplot_data_control <- convert_to_boxplot_data(clusters_ec_control)
boxplot_data_pq <- convert_to_boxplot_data(clusters_ec_pq)
boxplot_data_pq_RR <- convert_to_boxplot_data(clusters_ec_pq_RR)
boxplot_data_pq_gr <- convert_to_boxplot_data(clusters_ec_pq_gr)
boxplot_data_gr <- convert_to_boxplot_data(clusters_ec_gr)
boxplot_data_rr <- convert_to_boxplot_data(clusters_ec_rr)

# Summarize the data for each dataset
summarize_clusters <- function(boxplot_data) {
  boxplot_data %>%
    group_by(Cluster) %>%
    summarize(Average_EC = mean(EC), Min_EC = min(EC), Max_EC = max(EC))
}

# Create summary statistics for all datasets
cluster_summary_control <- summarize_clusters(boxplot_data_control)
cluster_summary_pq <- summarize_clusters(boxplot_data_pq)
cluster_summary_pq_RR <- summarize_clusters(boxplot_data_pq_RR)
cluster_summary_pq_gr <- summarize_clusters(boxplot_data_pq_gr)
cluster_summary_gr <- summarize_clusters(boxplot_data_gr)
cluster_summary_rr <- summarize_clusters(boxplot_data_rr)

# Combine all the summaries into one dataset
combined_summary <- bind_rows(
  mutate(cluster_summary_control, Dataset_Type = "Control"),
  mutate(cluster_summary_pq, Dataset_Type = "PQ"),
  mutate(cluster_summary_pq_RR, Dataset_Type = "PQ_RR"),
  mutate(cluster_summary_pq_gr, Dataset_Type = "PQ_GR"),
  mutate(cluster_summary_gr, Dataset_Type = "GR"),
  mutate(cluster_summary_rr, Dataset_Type = "RR")
)

#Ensure values are > 0 for log-log plot.
filtered_summary <- combined_summary %>%
  filter(Average_EC > 0)

#Uncomment if you want to "zoom in" on the first few clusters by removing values 
# below a certain tRReshold

# max_cluster_above_tRReshold <- filtered_summary %>%
#   summarise(max_cluster = max(Cluster)) %>%
#   pull(max_cluster)

# Filter the combined summary based on the maximum cluster index
#filtered_summary <- combined_summary %>%
#filter(Cluster <= max_cluster_above_tRReshold)

# Plot the trend lines with error bars for all groups using a log scale on the x-axis
ggplot(filtered_summary, aes(x = Cluster, y = Average_EC, color = Dataset_Type)) +
  geom_line(size = 1) +  
  geom_point(size = 2) +  
  #uncomment if you want to include error bars
  # geom_errorbar(aes(ymin = Min_EC, ymax = Max_EC), width = 0.01) +  
  labs(title = "Average EC per Cluster vs Mitochondrial Structure Index ",
       x = "Mitochondrial Structure Index", y = "Average Euler Characteristic") +
  scale_x_log10() +  
  scale_y_log10(limits = c(min(filtered_summary$Average_EC), max(filtered_summary$Average_EC))) +  
  theme_minimal() +
  theme(legend.position = "top")

###############
#Number of clusters Box-plot
###############

# Function to calculate the number of clusters per frame
average_clusters_per_dataset <- function(dataset) {
  clusters_per_frame <- colSums(!is.na(dataset))
    return(mean(clusters_per_frame))
}


con_avg_clusters <- sapply(datasetsCON, average_clusters_per_dataset)
pq_avg_clusters <- sapply(datasetsPQ, average_clusters_per_dataset)
pq_RR_avg_clusters <- sapply(datasetsPQ_RR, average_clusters_per_dataset)
pq_gr_avg_clusters <- sapply(datasetsPQ_GR, average_clusters_per_dataset)
gr_avg_clusters <- sapply(datasetsGR, average_clusters_per_dataset)
rr_avg_clusters <- sapply(datasetsRR, average_clusters_per_dataset)

cluster_data <- data.frame(
  Group = c(rep("CON", length(con_avg_clusters)),
            rep("PQ", length(pq_avg_clusters)),
            rep("GR", length(gr_avg_clusters)),
            rep("PQ+GR", length(pq_gr_avg_clusters)),
            rep("PQ+RR", length(pq_RR_avg_clusters)),
            rep("RR", length(rr_avg_clusters))),
  Avg_Clusters = c(con_avg_clusters, pq_avg_clusters, gr_avg_clusters, 
                   pq_gr_avg_clusters, pq_RR_avg_clusters, rr_avg_clusters)
)

ggplot(cluster_data, aes(x = Group, y = Avg_Clusters, fill = Group)) +
  geom_boxplot() +
  labs(title = "Average Number of Mitochondrial Structures per Group",
       x = "Group", y = "Average Number of Mitochondrial Structures") +
  theme_minimal() +
  theme(legend.position = "none")

# Remove outliers based on 1.5 IQR rule
remove_outliers <- function(data) {
  Q1 <- quantile(data, 0.25)
  Q3 <- quantile(data, 0.75)
  IQR <- Q3 - Q1
  lower_bound <- Q1 - 1.5 * IQR
  upper_bound <- Q3 + 1.5 * IQR
  data[data >= lower_bound & data <= upper_bound]
}

# Clean the data by removing outliers
clean_con<- remove_outliers(con_avg_clusters)
clean_pq<- remove_outliers(pq_avg_clusters)
clean_gr <- remove_outliers(gr_avg_clusters)
clean_rr <- remove_outliers(rr_avg_clusters)
clean_pq_gr <- remove_outliers(pq_RR_avg_clusters)
clean_pq_rr <- remove_outliers(pq_gr_avg_clusters)

t.test(clean_con, clean_pq)
t.test(clean_con, clean_gr)
t.test(clean_con, clean_rr)
t.test(clean_con, clean_pq_gr)
t.test(clean_con, clean_pq_rr)


#########
#Linear Regression Model
#########


# Apply log-transformation to the Cluster variable
filtered_summary$log_cluster <- log(filtered_summary$Cluster)
filtered_summary$log_Average_EC <- log(filtered_summary$Average_EC)

# Fit linear regression model with interaction for dataset type
linear_model <- lm(log_Average_EC ~ log_cluster * Dataset_Type, data = filtered_summary)

# View the summary of the linear regression model
summary(linear_model)

# Plot the residuals to check normality
par(mfrow = c(2, 2))  # Set up 2x2 plot grid
plot(linear_model)


ggplot(filtered_summary, aes(x = Average_EC)) + geom_histogram(bins = 100)


model_summary <- summary(linear_model)
coef_summary <- coef(model_summary)

power_law_results <- coef_summary[grep("log_cluster", rownames(coef_summary)), c("Estimate", "Std. Error")]
power_law_results <- as.data.frame(power_law_results)

# Clean up the row names to show only the Dataset_Type
rownames(power_law_results) <- sub("log_cluster:", "", rownames(power_law_results))

# Display the summary table
power_law_results



##############
#Ratio tests
##############

# Function to calculate the ratio of small (<2) to large (>=2) values
calculate_ratio <- function(dataset) {
  small_values <- sum(dataset <= 3, na.rm = TRUE)
  large_values <- sum(dataset > 3, na.rm = TRUE)
  
  if (large_values == 0) {
    return(NA)  # Avoid division by zero
  }
  
  ratio <- small_values / large_values
  return(ratio)
}

# Apply the ratio calculation to all datasets
ratiosCON <- sapply(datasetsCON, calculate_ratio)
ratiosPQ <- sapply(datasetsPQ, calculate_ratio)
ratiosPQ_RR <- sapply(datasetsPQ_RR, calculate_ratio)
ratiosPQ_GR <- sapply(datasetsPQ_GR, calculate_ratio)
ratiosGR <- sapply(datasetsGR, calculate_ratio)
ratiosRR <- sapply(datasetsRR, calculate_ratio)


all_ratios <- data.frame(
  Group = rep(c("CON", "PQ", "PQ_RR", "PQ_GR", "GR", "RR"), 
              times = c(length(ratiosCON), length(ratiosPQ), length(ratiosPQ_RR), 
                        length(ratiosPQ_GR), length(ratiosGR), length(ratiosRR))),
  Ratio = c(ratiosCON, ratiosPQ, ratiosPQ_RR, ratiosPQ_GR, ratiosGR, ratiosRR)
)

ggplot(all_ratios, aes(x = Group, y = Ratio, fill = Group)) +
  geom_boxplot() +
  theme_minimal() +
  labs(title = "Box Plot of Small to Large EC Ratios", y = "Small EC /Large EC Ratio", x = "Group")

t.test(ratiosCON, ratiosPQ)
t.test(ratiosCON, ratiosPQ_RR)
t.test(ratiosCON, ratiosPQ_GR)
t.test(ratiosCON, ratiosGR)
t.test(ratiosCON, ratiosRR)


#combine the small/large ratio with the average number of clusters
correlation_data <- data.frame(
  Group = cluster_data$Group,
  Avg_Clusters = cluster_data$Avg_Clusters,
  Small_Large_Ratio = c(ratiosCON, ratiosPQ, ratiosGR, 
                        ratiosPQ_GR, ratiosPQ_RR, ratiosRR)
)
print(correlation_data)
# Remove any rows with NA values that may exist due to division by zero in the ratios
correlation_data <- correlation_data[complete.cases(correlation_data), ]

# Scatter plot to visualize the relationship
ggplot(correlation_data, aes(x = Avg_Clusters, y = Small_Large_Ratio, color = Group)) +
  geom_point() +
  geom_smooth(method = "lm", se = FALSE) +  # Linear trend line
  labs(title = "Correlation between Small/Large Ratio and Average Clusters",
       x = "Average Number of Mitochondrial Structures",
       y = "Small EC /Large EC Ratio") +
  theme_minimal()

pearson_corr <- cor.test(correlation_data$Avg_Clusters, correlation_data$Small_Large_Ratio, method = "pearson")
print(pearson_corr)

spearman_corr <- cor.test(correlation_data$Avg_Clusters, correlation_data$Small_Large_Ratio, method = "spearman")
print(spearman_corr)


###########
# Average EC vs number of clusters
###########

# Function to calculate the average EC for cluster 1 in a dataset
calculate_avg_ec_cluster1 <- function(dataset) {
  # Extract the first EC value from each row (excluding the "Frame" label)
  first_ec_values <- as.numeric(dataset[1,])  # The second column holds the EC values for cluster 1
  
  # Calculate the average EC for cluster 1 over the frames
  avg_ec <- mean(first_ec_values, na.rm = TRUE)
  
  return(avg_ec)
}
# Function to calculate the number of clusters (mitochondria) in a dataset
calculate_avg_clusters <- function(dataset) {
  clusters_per_frame <- colSums(!is.na(dataset))  # Count non-NA values per frame
  avg_clusters <- mean(clusters_per_frame)  # Average clusters over all frames
  return(avg_clusters)
}

# Apply the functions to each dataset
con_ec <- sapply(datasetsCON, calculate_avg_ec_cluster1)
pq_ec <- sapply(datasetsPQ, calculate_avg_ec_cluster1)
pq_RR_ec <- sapply(datasetsPQ_RR, calculate_avg_ec_cluster1)
pq_gr_ec <- sapply(datasetsPQ_GR, calculate_avg_ec_cluster1)
gr_ec <- sapply(datasetsGR, calculate_avg_ec_cluster1)
rr_ec <- sapply(datasetsRR, calculate_avg_ec_cluster1)

con_clusters <- sapply(datasetsCON, calculate_avg_clusters)
pq_clusters <- sapply(datasetsPQ, calculate_avg_clusters)
pq_RR_clusters <- sapply(datasetsPQ_RR, calculate_avg_clusters)
pq_gr_clusters <- sapply(datasetsPQ_GR, calculate_avg_clusters)
gr_clusters <- sapply(datasetsGR, calculate_avg_clusters)
rr_clusters <- sapply(datasetsRR, calculate_avg_clusters)

# Combine the EC and cluster data into a data frame
ec_cluster_data <- data.frame(
  Group = c(rep("CON", length(con_ec)),
            rep("PQ", length(pq_ec)),
            rep("PQ_RR", length(pq_RR_ec)),
            rep("PQ_GR", length(pq_gr_ec)),
            rep("GR", length(gr_ec)),
            rep("RR", length(rr_ec))),
  
  Avg_EC_Cluster1 = c(con_ec, pq_ec, pq_RR_ec, pq_gr_ec, gr_ec, rr_ec),
  Avg_Clusters = c(con_clusters, pq_clusters, pq_RR_clusters, pq_gr_clusters, gr_clusters, rr_clusters)
)

# Plot the relationship between Avg EC at Cluster 1 and Avg Clusters
ggplot(ec_cluster_data, aes(x = Avg_Clusters, y = Avg_EC_Cluster1, color = Group)) +
  geom_point() +
  geom_smooth(method = "lm", se = FALSE) +  # Add linear trend line
  labs(title = "Correlation between Avg EC at Mitochondrial Index 1 and Number of Mitochondria",
       x = "Average Number of Mitochondrial Structures",
       y = "Average EC at Mitochondrial Index 1") +
  theme_minimal()
