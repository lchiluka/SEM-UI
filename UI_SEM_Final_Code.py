import streamlit as st
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import os
import time
import cv2
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy import ndimage as ndi
from concurrent.futures import ThreadPoolExecutor, as_completed
import shutil

# SharePoint details
sharepoint_base_url = 'https://carlislecompanies.sharepoint.com/teams/CCMRD7857'
folder_in_sharepoint = '/teams/CCMRD7857/RDrive/Analytical/2024%20Projects/Insulation'

# Function to log messages
def log(message):
    """Logs a message to the Streamlit interface."""
    st.write(message)

if "overall_execution_time" not in st.session_state:
    st.session_state.overall_execution_time = 0
if "overall_files_processed" not in st.session_state:
    st.session_state.overall_files_processed = 0
if "unprocessed_images" not in st.session_state:
    st.session_state.unprocessed_images = []

# Function to authenticate to SharePoint
def authenticate_to_sharepoint(username, password):
    """
    Authenticates a user to SharePoint using the provided username and password.
    
    Args:
        username (str): The SharePoint username.
        password (str): The SharePoint password.

    Returns:
        ClientContext: The authenticated SharePoint client context, or None if authentication fails.
    """
    try:
        auth_ctx = AuthenticationContext(sharepoint_base_url)
        if auth_ctx.acquire_token_for_user(username, password):
            ctx = ClientContext(sharepoint_base_url, auth_ctx)
            web = ctx.web
            ctx.load(web)
            ctx.execute_query()
            return ctx
        else:
            st.error("Authentication Failed: Invalid username or password")
            return None
    except Exception as e:
        st.error(f"Invalid username or password")
        return None

# Function to get available folders from SharePoint
def get_available_folders(ctx, folder_url):
    """
    Retrieves a list of available folders from a specified SharePoint folder URL.
    
    Args:
        ctx (ClientContext): The SharePoint client context.
        folder_url (str): The relative URL of the SharePoint folder.

    Returns:
        list: A list of folder names within the specified SharePoint folder.
    """
    try:
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        folder.expand(["Folders"]).get().execute_query()
        folder_names = [f.properties["Name"] for f in folder.folders]
        return folder_names
    except Exception as e:
        log(f"Failed to retrieve folders: {e}")
        return []

# Function to get PNG files in a SharePoint folder
def get_png_files_in_folder(ctx, folder_url):
    """
    Retrieves a list of PNG files from a specified SharePoint folder URL.
    
    Args:
        ctx (ClientContext): The SharePoint client context.
        folder_url (str): The relative URL of the SharePoint folder.

    Returns:
        list: A list of tuples containing the folder URL and file name for each PNG file.
    """
    png_files = []
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    folder.expand(["Files", "Folders"]).get().execute_query()
    for file in folder.files:
        file_name = file.properties['Name']
        if file_name.lower().endswith('.png'):
            png_files.append((folder_url, file_name))
    return png_files

# Function to create a folder in SharePoint
def create_folder(ctx, parent_folder_url, new_folder_name, max_retries=5, initial_wait=1):
    """
    Creates a new folder in SharePoint under the specified parent folder.
    
    Args:
        ctx (ClientContext): The SharePoint client context.
        parent_folder_url (str): The relative URL of the parent folder in SharePoint.
        new_folder_name (str): The name of the new folder to create.
        max_retries (int): The maximum number of retry attempts in case of failure.
        initial_wait (int): The initial wait time in seconds between retries.

    Returns:
        None
    """
    attempt = 0
    while attempt < max_retries:
        try:
            parent_folder = ctx.web.get_folder_by_server_relative_url(parent_folder_url)
            ctx.load(parent_folder)
            ctx.execute_query()  # Ensure parent folder exists
            
            new_folder = parent_folder.folders.add(new_folder_name)
            ctx.execute_query()  # Create new folder
            print(f"Folder '{new_folder_name}' created successfully at '{parent_folder_url}'")
            return
        except Exception as e:
            attempt += 1
            wait_time = initial_wait * (2 ** attempt)  # Exponential backoff
            log(f"Error creating folder '{new_folder_name}' at '{parent_folder_url}': {e}. Retrying in {wait_time} seconds...")
            time.sleep(wait_time)

    # If all retries failed, log the failure
    log(f"Failed to create folder '{new_folder_name}' at '{parent_folder_url}' after {max_retries} attempts.")
    st.session_state.unprocessed_images.append(f"{parent_folder_url}/{new_folder_name}")

# Function to download a file from SharePoint
def download_file(ctx, file_relative_url, local_path):
    """
    Downloads a file from SharePoint to a local path.
    
    Args:
        ctx (ClientContext): The SharePoint client context.
        file_relative_url (str): The relative URL of the file in SharePoint.
        local_path (str): The local path where the file will be saved.

    Returns:
        None
    """
    try:
        response = File.open_binary(ctx, file_relative_url)
        with open(local_path, "wb") as local_file:
            local_file.write(response.content)
        print(f"Downloaded: {local_path}")
    except Exception as e:
        log(f"Error downloading {file_relative_url}: {e}")
        st.session_state.unprocessed_images.append(file_relative_url)

# Function to upload a file to SharePoint
def upload_file(ctx, target_folder_url, file_name, local_path):
    """
    Uploads a file from a local path to a specified SharePoint folder.
    
    Args:
        ctx (ClientContext): The SharePoint client context.
        target_folder_url (str): The relative URL of the target folder in SharePoint.
        file_name (str): The name of the file to upload.
        local_path (str): The local path of the file to upload.

    Returns:
        None
    """
    try:
        with open(local_path, "rb") as local_file:
            file_content = local_file.read()
        target_folder = ctx.web.get_folder_by_server_relative_url(target_folder_url)
        target_file = target_folder.upload_file(file_name, file_content).execute_query()
        print(f"Uploaded: {file_name} to {target_folder_url}")
    except Exception as e:
        log(f"Error uploading {file_name}: {e}")
        st.session_state.unprocessed_images.append(local_path)

def calculate_acs(image, pixel_to_um=2):
    """
    Calculates the Average Cell Size (ACS) in vertical and horizontal directions.
    
    Args:
        image (numpy.ndarray): The binary image used to calculate ACS.
        pixel_to_um (int, optional): The conversion factor from pixels to micrometers. Default is 2.

    Returns:
        tuple: The average vertical ACS and average horizontal ACS in micrometers.
    """
    height, width = image.shape
    vertical_lines = np.linspace(0, width - 1, num=height // 20, dtype=int)
    horizontal_lines = np.linspace(0, height - 1, num=width // 20, dtype=int)

    vertical_acs = []
    for x in vertical_lines:
        column = image[:, x]
        edges = np.diff((column > 0).astype(int))
        num_cells = np.sum(edges == 1)
        acs = height / num_cells if num_cells > 0 else 0
        vertical_acs.append(acs)

    horizontal_acs = []
    for y in horizontal_lines:
        row = image[y, :]
        edges = np.diff((row > 0).astype(int))
        num_cells = np.sum(edges == 1)
        acs = width / num_cells if num_cells > 0 else 0
        horizontal_acs.append(acs)

    avg_vertical_acs = np.mean(vertical_acs) * pixel_to_um
    avg_horizontal_acs = np.mean(horizontal_acs) * pixel_to_um

    return avg_vertical_acs, avg_horizontal_acs

def calculate_ar(vertical_acs, horizontal_acs):
    """
    Calculates the Anisotropy Ratio (AR) based on vertical and horizontal ACS.
    
    Args:
        vertical_acs (float): The average vertical ACS.
        horizontal_acs (float): The average horizontal ACS.

    Returns:
        float: The Anisotropy Ratio (AR).
    """
    ar = vertical_acs / horizontal_acs if horizontal_acs > 0 else 0
    return ar

def process_image(image_path, pixel_to_um=2):
    """
    Processes an image to extract cellular properties and calculates various metrics.
    
    Args:
        image_path (str): The path to the image file.
        pixel_to_um (int, optional): The conversion factor from pixels to micrometers. Default is 2.

    Returns:
        tuple: A tuple containing the following elements:
            - regions_properties (list): A list of dictionaries with the properties of each region.
            - solid_contours (list): A list of solid contours found in the image.
            - color_image (numpy.ndarray): The processed color image.
            - avg_vertical_acs (float): The average vertical ACS.
            - avg_horizontal_acs (float): The average horizontal ACS.
            - ar (float): The Anisotropy Ratio (AR).
    """
    try:
        image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
        blurred = cv2.GaussianBlur(image, (15, 15), 0)
        high_pass_kernel = np.array([[-1, -1, -1],
                                     [-1,  8, -1],
                                     [-1, -1, -1]])
        high_pass_filtered = cv2.filter2D(blurred, -2, high_pass_kernel)
        high_pass_filtered = cv2.normalize(high_pass_filtered, None, 0, 255, cv2.NORM_MINMAX)

        _, binary = cv2.threshold(high_pass_filtered, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        dist_transform = cv2.distanceTransform(binary, cv2.DIST_L2, 5)
        _, markers = cv2.threshold(dist_transform, 0.3 * dist_transform.max(), 255, 0)
        markers = np.uint8(markers)

        sure_bg = cv2.dilate(binary, np.ones((3, 3), np.uint8), iterations=5)
        sure_fg = cv2.erode(binary, np.ones((3, 3), np.uint8), iterations=5)
        unknown = cv2.subtract(sure_bg, sure_fg)

        _, markers = cv2.connectedComponents(sure_fg)
        markers = markers + 1
        markers[unknown == 255] = 0

        color_image = cv2.cvtColor(high_pass_filtered, cv2.COLOR_GRAY2BGR)
        cv2.watershed(color_image, markers)
        color_image[markers == -1] = [0, 255, 0]

        gray_color_image = cv2.cvtColor(color_image, cv2.COLOR_BGR2GRAY)
        blurred = cv2.GaussianBlur(gray_color_image, (5, 5), 0)
        binary = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2)

        binary[:40, :] = 0
        binary[-40:, :] = 0
        binary[:, :20] = 0
        binary[:, -20:] = 0

        num_labels, labels = cv2.connectedComponents(binary)
        area_threshold = 70

        filtered_labels = np.zeros_like(labels)
        for i in range(1, num_labels):
            cell_mask = labels == i
            if np.sum(cell_mask) >= area_threshold:
                filtered_labels[cell_mask] = i

        split_labels = split_cells(filtered_labels, area_threshold)

        for i in range(1, split_labels.max() + 1):
            cell_mask = split_labels == i
            if np.any(cell_mask):
                filled_cell = ndi.binary_fill_holes(cell_mask).astype(int)
                split_labels[filled_cell > 0] = i

        colored_cells = np.zeros((split_labels.shape[0], split_labels.shape[1], 3), dtype=np.uint8)
        for i in range(1, split_labels.max() + 1):
            cell_mask = split_labels == i
            if np.any(cell_mask):
                contours, _ = cv2.findContours(cell_mask.astype(np.uint8), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                color = np.random.randint(0, 255, size=3).tolist()
                cv2.drawContours(colored_cells, contours, -1, color, -1)

        gray = cv2.cvtColor(colored_cells, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 1, 255, cv2.THRESH_BINARY)

        contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        solid_contours = [cnt for cnt in contours if is_solid_region(cnt)]

        regions_properties = []
        for cnt in solid_contours:
            area = cv2.contourArea(cnt) * (pixel_to_um ** 2)
            perimeter = cv2.arcLength(cnt, True) * pixel_to_um
            equivalent_diameter = np.sqrt(4 * area / np.pi)
            area_x_equivalent_diameter = area * equivalent_diameter
            x, y, w, h = cv2.boundingRect(cnt)
            min_diameter = min(w, h) * pixel_to_um
            max_diameter = max(w, h) * pixel_to_um
            mean_diameter = np.mean([w, h]) * pixel_to_um
            aspect_ratio = w / h
            elongation = abs(1 - (w / h))
            if len(cnt) >= 5:
                ellipse = cv2.fitEllipse(cnt)
                angle = ellipse[2]
                # Instead of relying on the default angle, we manually flip the angle.
                orientation = (90 - angle) if angle < 90 else (270 - angle)
            else:
                orientation = np.nan
            roundness = (4 * area) / (np.pi * (mean_diameter ** 2))
            centroid = (x + w // 2, y + h // 2)
            regions_properties.append({
                'Projected area (µm²)': area,
                'Perimeter (µm)': perimeter,
                'Equivalent diameter (µm)': equivalent_diameter,
                'Area x equivalent diameter (µm)': area_x_equivalent_diameter,
                'Mean diameter (µm)': mean_diameter,
                'Minimum diameter (µm)': min_diameter,
                'Maximum diameter (µm)': max_diameter,
                'Length (µm)': h * pixel_to_um,
                'Width (µm)': w * pixel_to_um,
                'Aspect ratio': aspect_ratio,
                'Roundness': roundness,
                'Elongation': elongation,
                'Orientation (°)': orientation,
                'Centroid': centroid
            })

        avg_vertical_acs, avg_horizontal_acs = calculate_acs(binary, pixel_to_um)
        ar = calculate_ar(avg_vertical_acs, avg_horizontal_acs)

        return regions_properties, solid_contours, color_image, avg_vertical_acs, avg_horizontal_acs, ar
    except Exception as e:
        log(f"Error processing {image_path}: {e}")
        return [], [], None, 0, 0, 0

def split_cells(labels, area_threshold):
    """
    Splits connected cells in a labeled image using distance transform and connected components analysis.
    
    Args:
        labels (numpy.ndarray): The labeled image.
        area_threshold (int): The minimum area threshold to consider a cell for splitting.

    Returns:
        numpy.ndarray: The re-labeled image after splitting cells.
    """
    new_labels = np.zeros_like(labels)
    current_label = 1
    for i in range(1, labels.max() + 1):
        cell_mask = labels == i
        if np.sum(cell_mask) >= area_threshold:
            dist_transform = cv2.distanceTransform(cell_mask.astype(np.uint8), cv2.DIST_L2, 5)
            _, split_markers = cv2.threshold(dist_transform, 0.2 * dist_transform.max(), 255, 0)
            split_markers = np.uint8(split_markers)
            num_splits, split_labels = cv2.connectedComponents(split_markers)
            for j in range(1, num_splits):
                split_mask = split_labels == j
                if np.sum(split_mask) >= area_threshold:
                    new_labels[split_mask] = current_label
                    current_label += 1
        else:
            new_labels[cell_mask] = current_label
            current_label += 1
    return new_labels

def is_solid_region(contour):
    """
    Determines if a contour represents a solid region based on its area and perimeter.
    
    Args:
        contour (numpy.ndarray): The contour to analyze.

    Returns:
        bool: True if the region is considered solid, False otherwise.
    """
    area = cv2.contourArea(contour)
    perimeter = cv2.arcLength(contour, True)
    return area > 100 and perimeter / area < 0.1

def calculate_statistics(df, avg_vertical_acs, avg_horizontal_acs, ar):
    """
    Calculates statistical metrics for a DataFrame of region properties and adds additional calculated metrics.
    
    Args:
        df (pandas.DataFrame): The DataFrame containing region properties.
        avg_vertical_acs (float): The average vertical ACS.
        avg_horizontal_acs (float): The average horizontal ACS.
        ar (float): The Anisotropy Ratio (AR).

    Returns:
        pandas.DataFrame: A DataFrame containing calculated statistics.
    """
    stats = df.describe().T
    stats['median'] = df.median()
    stats = stats[['mean', 'std', 'min', 'max', 'median']]
    stats = stats.T
    stats['Statistic'] = stats.index
    stats = stats.reset_index(drop=True)

    total_area = df['Projected area (µm²)'].sum()
    if total_area > 0:
        area_weighted_equivalent_diameter = df['Area x equivalent diameter (µm)'].sum() / total_area
        stats['Area weighted equivalent diameter (µm)'] = [area_weighted_equivalent_diameter, np.nan, np.nan, np.nan, np.nan]

    stats['Anisotropy Ratio (AR)'] = [ar, np.nan, np.nan, np.nan, np.nan]

    return stats

def create_image2_format_table(df, avg_vertical_acs, avg_horizontal_acs, ar):
    """
    Creates a table with averages and standard deviations, similar to the second image format.
    
    Args:
        df (pandas.DataFrame): The DataFrame containing region properties.
        avg_vertical_acs (float): The average vertical ACS.
        avg_horizontal_acs (float): The average horizontal ACS.
        ar (float): The Anisotropy Ratio (AR).
    
    Returns:
        pandas.DataFrame: A DataFrame containing averages and standard deviations in the required format.
    """
    # Create a DataFrame for averages and standard deviations
    stats_image2 = df.describe().T[['mean', 'std']]
    
    # Renaming the columns for clarity
    stats_image2.columns = ['Average', 'STD']
    
    # Round the values to 2 decimal places
    stats_image2 = stats_image2.round(2)
    
    # Add additional calculated metrics (e.g., ACS and AR)
    total_area = df['Projected area (µm²)'].sum()
    if total_area > 0:
        area_weighted_equivalent_diameter = df['Area x equivalent diameter (µm)'].sum() / total_area
        stats_image2.loc['Area weighted equivalent diameter (µm)', 'Average'] = round(area_weighted_equivalent_diameter, 2)
    
    # Adding vertical, horizontal ACS, and AR
    stats_image2.loc['Average vertical ACS (µm)', 'Average'] = round(avg_vertical_acs, 2)
    stats_image2.loc['Average horizontal ACS (µm)', 'Average'] = round(avg_horizontal_acs, 2)
    stats_image2.loc['Anisotropy Ratio (AR)', 'Average'] = round(ar, 2)

    # Drop 'Area x equivalent diameter (µm)'
    if 'Area x equivalent diameter (µm)' in stats_image2.index:
        stats_image2 = stats_image2.drop('Area x equivalent diameter (µm)')

    # Add the row names and reset index for display purposes
    stats_image2 = stats_image2.reset_index().rename(columns={'index': 'Metric'})

    return stats_image2


def plot_and_save_image(image, contours, centroids, output_path):
    """
    Plots and saves an image with overlaid contours and centroids.
    
    Args:
        image (numpy.ndarray): The image to plot.
        contours (list): A list of contours to overlay on the image.
        centroids (list): A list of centroids to overlay on the image.
        output_path (str): The path where the image will be saved.

    Returns:
        None
    """
    plt.figure(figsize=(10, 10))
    plt.imshow(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))
    
    for contour in contours:
        plt.plot(contour[:, 0, 0], contour[:, 0, 1], 'b-', linewidth=1)
    
    for centroid in centroids:
        plt.plot(centroid[0], centroid[1], 'r.', markersize=10)
    
    plt.axis('off')
    plt.savefig(output_path, bbox_inches='tight')
    plt.close()

def process_file(root, file, base_dir, export_dir, pixel_to_um=2):
    """
    Processes a single image file to extract and save region properties and statistics.
    
    Args:
        root (str): The root directory of the image file.
        file (str): The name of the image file.
        base_dir (str): The base directory for processing.
        export_dir (str): The directory where processed files will be saved.
        pixel_to_um (int, optional): The conversion factor from pixels to micrometers. Default is 2.

    Returns:
        tuple: A tuple containing the subfolder key, file name, DataFrame of properties, average vertical ACS, 
               average horizontal ACS, and Anisotropy Ratio (AR).
    """
    try:
        image_path = os.path.join(root, file)
        properties, contours, processed_image, avg_vertical_acs, avg_horizontal_acs, ar = process_image(image_path, pixel_to_um)
        properties_df = pd.DataFrame(properties)

        output_subfolder_key = os.path.relpath(root, base_dir)
        output_image_path = os.path.join(export_dir, output_subfolder_key, f"{os.path.splitext(file)[0]}_processed.png")
        os.makedirs(os.path.dirname(output_image_path), exist_ok=True)
        centroids = [prop['Centroid'] for prop in properties]
        plot_and_save_image(processed_image, contours, centroids, output_image_path)

        return output_subfolder_key, file, properties_df, avg_vertical_acs, avg_horizontal_acs, ar
    except Exception as e:
        log(f"Error processing file {file} in {root}: {e}")
        return None

def traverse_and_process(base_dir, export_dir, max_files=2, pixel_to_um=2):
    """
    Traverses directories and processes image files to extract region properties, 
    calculate statistics, and save results.
    
    Args:
        base_dir (str): The base directory to traverse.
        export_dir (str): The directory where processed files will be saved.
        max_files (int, optional): The maximum number of files to process in each folder. Default is 2.
        pixel_to_um (int, optional): The conversion factor from pixels to micrometers. Default is 2.

    Returns:
        None
    """
    start_time = time.time()
    file_count = 0

    # Dictionary to store results from all folders
    subfolder_dfs = {}

    with ThreadPoolExecutor(max_workers=24) as executor:
        futures = []

        for root, dirs, files in os.walk(base_dir):
            if 'Exported Results' in root:
                continue

            # Process all PNG files, not just 'MD' or 'TD'
            png_files = [f for f in files if f.endswith('.PNG')][:max_files]

            # Submit the processing of each PNG file to the executor pool
            for file in png_files:
                futures.append(executor.submit(process_file, root, file, base_dir, export_dir, pixel_to_um))

        # Collect results as they are completed
        for future in as_completed(futures):
            result = future.result()
            if result:
                subfolder_key, file, properties_df, avg_vertical_acs, avg_horizontal_acs, ar = result

                # Store results in the dictionary for each folder
                if subfolder_key not in subfolder_dfs:
                    subfolder_dfs[subfolder_key] = properties_df
                else:
                    subfolder_dfs[subfolder_key] = pd.concat([subfolder_dfs[subfolder_key], properties_df], ignore_index=True)

                # Add calculated metrics to the DataFrame
                subfolder_dfs[subfolder_key]['Average vertical ACS (µm)'] = avg_vertical_acs
                subfolder_dfs[subfolder_key]['Average horizontal ACS (µm)'] = avg_horizontal_acs
                subfolder_dfs[subfolder_key]['Anisotropy Ratio (AR)'] = ar

                file_count += 1

    # Save aggregated results to Excel files for each folder
    for subfolder_key, combined_df in subfolder_dfs.items():
        output_subfolder = os.path.join(export_dir, subfolder_key)
        os.makedirs(output_subfolder, exist_ok=True)

        if 'Centroid' in combined_df.columns:
            combined_df = combined_df.drop(columns=['Centroid'])

        output_file = os.path.join(output_subfolder, f"{os.path.basename(subfolder_key)}_aggregated.xlsx")

        avg_vertical_acs = combined_df['Average vertical ACS (µm)'].iloc[0]
        avg_horizontal_acs = combined_df['Average horizontal ACS (µm)'].iloc[0]
        ar = combined_df['Anisotropy Ratio (AR)'].iloc[0]

        # Calculate original statistics
        original_stats_df = calculate_statistics(combined_df, avg_vertical_acs, avg_horizontal_acs, ar)
        # Create the new table in the format similar to Image 2
        image2_table_df = create_image2_format_table(combined_df, avg_vertical_acs, avg_horizontal_acs, ar)

        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            combined_df.to_excel(writer, sheet_name='Data', index=False)
            original_stats_df.to_excel(writer, sheet_name='Original Statistics', index=False)
            image2_table_df.to_excel(writer, sheet_name='Image2 Format', index=False)

            # Access the workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Image2 Format']
        
            # Define formatting for specific cells
            yellow_format = workbook.add_format({'bg_color': '#FFFF00'})  # Yellow
            blue_format = workbook.add_format({'bg_color': '#00BFFF'})    # Blue
            orange_format = workbook.add_format({'bg_color': '#FFA500'})  # Orange
            green_format = workbook.add_format({'bg_color': '#00FF00'})   # Green

            # Apply conditional formatting for specific rows and metrics
            for row_num, metric in enumerate(image2_table_df['Metric'], start=1):
                if metric == 'Mean diameter (µm)':
                    worksheet.write(row_num, 1, image2_table_df.loc[row_num - 1, 'Average'], yellow_format)
                elif metric == 'Minimum diameter (µm)' or metric == 'Maximum diameter (µm)':
                    worksheet.write(row_num, 1, image2_table_df.loc[row_num - 1, 'Average'], blue_format)
                elif metric == 'Orientation (°)':
                    worksheet.write(row_num, 1, image2_table_df.loc[row_num - 1, 'Average'], orange_format)
                elif metric == 'Area weighted equivalent diameter (µm)':
                    worksheet.write(row_num, 1, image2_table_df.loc[row_num - 1, 'Average'], green_format)

    end_time = time.time()
    execution_time = end_time - start_time
    st.session_state.overall_execution_time += execution_time
    st.session_state.overall_files_processed += file_count
    log(f"Total execution time: {execution_time:.2f} seconds")
    log(f"Total files processed: {file_count}")


def process_folder_and_subfolders(ctx, folder_url):
    """
    Recursively processes folders and subfolders in SharePoint, downloads image files, processes them locally,
    and uploads the results back to SharePoint.
    
    Args:
        ctx (ClientContext): The SharePoint client context.
        folder_url (str): The relative URL of the folder to process in SharePoint.

    Returns:
        None
    """
    folder_name = folder_url.split('/')[-1]

    # Skip folders that start with 'output_'
    if folder_name.startswith('final_output_'):
        return

    # Process folders named 'core', 'top', or 'bottom', and their subfolders
    log(f"Processing Folder: {folder_name}")
    png_files_list = get_png_files_in_folder(ctx, folder_url)

    if png_files_list:
        # Create local folders for download and export
        local_output_folder = f"output_{folder_name}"
        local_export_folder = os.path.join(local_output_folder, 'export')
        os.makedirs(local_output_folder, exist_ok=True)

        # Download each .png file in the list
        for folder_path, file_name in png_files_list:
            file_relative_url = f"{folder_path}/{file_name}"
            local_file_path = os.path.join(local_output_folder, file_name)
            download_file(ctx, file_relative_url, local_file_path)

        # Process the images
        traverse_and_process(local_output_folder, local_export_folder, max_files=50)

        # Upload processed files back to SharePoint
        export_folder_in_sharepoint = f"{folder_url}/final_output_{folder_name}"
        create_folder(ctx, folder_url, f"final_output_{folder_name}")

        for file_name in os.listdir(local_export_folder):
            local_file_path = os.path.join(local_export_folder, file_name)
            upload_file(ctx, export_folder_in_sharepoint, file_name, local_file_path)

        # Delete the local downloaded folder
        shutil.rmtree(local_output_folder)
        log(f"Deleted local folder: {local_output_folder}")

    # Recurse into all subfolders
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    folder.expand(["Folders"]).get().execute_query()

    for subfolder in folder.folders:
        subfolder_url = f"{folder_url}/{subfolder.properties['Name']}"
        process_folder_and_subfolders(ctx, subfolder_url)


st.image("Carlisle_MasterLogo_RGB.jpg", width=500)  # Adjust the width as needed
st.title("Select the folder in SharePoint for SEM Analysis")

# Input fields for SharePoint username and password
username = st.text_input("SharePoint Username", key="username")
password = st.text_input("SharePoint Password", type="password", key="password")

# Authenticate button
if st.button("Authenticate"):
    ctx = authenticate_to_sharepoint(username, password)
    if ctx:
        st.success("Authenticated Successfully")
        st.session_state.ctx = ctx

        # Get available folders
        folders = ["Select"] + get_available_folders(ctx, folder_in_sharepoint)
        st.session_state.folders = folders
        
if "folders" in st.session_state:
    # Dropdown to select a folder
    selected_folder = st.selectbox("Select a Folder", st.session_state.folders)
    
    if selected_folder != "Select":
        # Store the selected folder path
        selected_folder_path = f"{folder_in_sharepoint}/{selected_folder}"
        st.write(f"Selected Folder Path: {selected_folder_path}")
        
        # Generate button
        if st.button("Generate", key="generate_button"):
            process_folder_and_subfolders(st.session_state.ctx, selected_folder_path)
            st.success("Processing Completed Successfully")
            log(f"Overall execution time: {st.session_state.overall_execution_time:.2f} seconds")
            log(f"Overall files processed: {st.session_state.overall_files_processed}")
            log("Failed files:")
            for image in st.session_state.unprocessed_images:
                log(image)
