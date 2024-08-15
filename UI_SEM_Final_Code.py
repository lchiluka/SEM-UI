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
    st.write(message)

if "overall_execution_time" not in st.session_state:
    st.session_state.overall_execution_time = 0
if "overall_files_processed" not in st.session_state:
    st.session_state.overall_files_processed = 0
if "unprocessed_images" not in st.session_state:
    st.session_state.unprocessed_images = []

# Function to authenticate to SharePoint
def authenticate_to_sharepoint(username, password):
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
    ar = vertical_acs / horizontal_acs if horizontal_acs > 0 else 0
    return ar

def process_image(image_path, pixel_to_um=2):
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
                (major_axis, minor_axis), angle = ellipse[1], ellipse[2]
                orientation = angle
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
    area = cv2.contourArea(contour)
    perimeter = cv2.arcLength(contour, True)
    return area > 100 and perimeter / area < 0.1

def calculate_statistics(df, avg_vertical_acs, avg_horizontal_acs, ar):
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

def plot_and_save_image(image, contours, centroids, output_path):
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
    start_time = time.time()
    file_count = 0

    subfolder_md_dfs = {}
    subfolder_td_dfs = {}

    with ThreadPoolExecutor(max_workers=24) as executor:
        futures = []

        for root, dirs, files in os.walk(base_dir):
            if 'Exported Results' in root:
                continue

            if os.path.basename(base_dir) not in root and base_dir not in root:
                continue

            md_files = [f for f in files if f.endswith('.PNG') and 'MD' in f][:max_files]
            td_files = [f for f in files if f.endswith('.PNG') and 'TD' in f][:max_files]

            for file in md_files:
                futures.append(executor.submit(process_file, root, file, base_dir, export_dir, pixel_to_um))

            for file in td_files:
                futures.append(executor.submit(process_file, root, file, base_dir, export_dir, pixel_to_um))

        for future in as_completed(futures):
            result = future.result()
            if result:
                subfolder_key, file, properties_df, avg_vertical_acs, avg_horizontal_acs, ar = result
                if 'MD' in file:
                    if subfolder_key not in subfolder_md_dfs:
                        subfolder_md_dfs[subfolder_key] = properties_df
                    else:
                        subfolder_md_dfs[subfolder_key] = pd.concat([subfolder_md_dfs[subfolder_key], properties_df], ignore_index=True)
                    subfolder_md_dfs[subfolder_key]['Average vertical ACS (µm)'] = avg_vertical_acs
                    subfolder_md_dfs[subfolder_key]['Average horizontal ACS (µm)'] = avg_horizontal_acs
                    subfolder_md_dfs[subfolder_key]['Anisotropy Ratio (AR)'] = ar
                else:
                    if subfolder_key not in subfolder_td_dfs:
                        subfolder_td_dfs[subfolder_key] = properties_df
                    else:
                        subfolder_td_dfs[subfolder_key] = pd.concat([subfolder_td_dfs[subfolder_key], properties_df], ignore_index=True)
                    subfolder_td_dfs[subfolder_key]['Average vertical ACS (µm)'] = avg_vertical_acs
                    subfolder_td_dfs[subfolder_key]['Average horizontal ACS (µm)'] = avg_horizontal_acs
                    subfolder_td_dfs[subfolder_key]['Anisotropy Ratio (AR)'] = ar

                file_count += 1
                #log(f"Processed file: {file}")

    for subfolder_key, combined_md_df in subfolder_md_dfs.items():
        output_subfolder = os.path.join(export_dir, subfolder_key)
        os.makedirs(output_subfolder, exist_ok=True)

        if 'Centroid' in combined_md_df.columns:
            combined_md_df = combined_md_df.drop(columns=['Centroid'])

        output_file_md = os.path.join(output_subfolder, f"{os.path.basename(subfolder_key)} MD.xlsx")

        avg_vertical_acs = combined_md_df['Average vertical ACS (µm)'].iloc[0]
        avg_horizontal_acs = combined_md_df['Average horizontal ACS (µm)'].iloc[0]
        ar = combined_md_df['Anisotropy Ratio (AR)'].iloc[0]

        stats_md_df = calculate_statistics(combined_md_df, avg_vertical_acs, avg_horizontal_acs, ar)

        with pd.ExcelWriter(output_file_md) as writer:
            combined_md_df.to_excel(writer, sheet_name='Data', index=False)
            stats_md_df.to_excel(writer, sheet_name='Statistics', index=False)

    for subfolder_key, combined_td_df in subfolder_td_dfs.items():
        output_subfolder = os.path.join(export_dir, subfolder_key)
        os.makedirs(output_subfolder, exist_ok=True)

        if 'Centroid' in combined_td_df.columns:
            combined_td_df = combined_td_df.drop(columns=['Centroid'])

        output_file_td = os.path.join(output_subfolder, f"{os.path.basename(subfolder_key)} TD.xlsx")

        avg_vertical_acs = combined_td_df['Average vertical ACS (µm)'].iloc[0]
        avg_horizontal_acs = combined_td_df['Average horizontal ACS (µm)'].iloc[0]
        ar = combined_td_df['Anisotropy Ratio (AR)'].iloc[0]

        if not combined_td_df.empty:
            stats_td_df = calculate_statistics(combined_td_df, avg_vertical_acs, avg_horizontal_acs, ar)

            with pd.ExcelWriter(output_file_td) as writer:
                combined_td_df.to_excel(writer, sheet_name='Data', index=False)
                stats_td_df.to_excel(writer, sheet_name='Statistics', index=False)

    end_time = time.time()
    execution_time = end_time - start_time
    st.session_state.overall_execution_time += execution_time
    st.session_state.overall_files_processed += file_count
    log(f"Total execution time: {execution_time:.2f} seconds")
    log(f"Total files processed: {file_count}")

def process_folder_and_subfolders(ctx, folder_url):
    folder_name = folder_url.split('/')[-1]

    # Skip folders that start with 'output_'
    if folder_name.startswith('final_output_'):
        return

    # Only process folders named 'top', 'bottom', or 'core'
    log(f"Folder Name: {folder_name}")
    if folder_name.lower() in ['top', 'bottom', 'core']:
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
            export_folder_in_sharepoint = f"{folder_url}/ultimate_output_{folder_name}"
            create_folder(ctx, folder_url, f"ultimate_output_{folder_name}")

            for file_name in os.listdir(local_export_folder):
                local_file_path = os.path.join(local_export_folder, file_name)
                upload_file(ctx, export_folder_in_sharepoint, file_name, local_file_path)

            # Delete the local downloaded folder
            shutil.rmtree(local_output_folder)
            print(f"Deleted local folder: {local_output_folder}")

    # Recurse into subfolders
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
