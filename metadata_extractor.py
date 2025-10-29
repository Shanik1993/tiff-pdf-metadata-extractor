import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from PIL import Image, TiffTags
import pandas as pd
import re
import PyPDF2
from PyPDF2 import PdfReader
import io

class MetadataExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("TIFF and PDF Metadata Extractor")
        self.root.geometry("1200x700")
        
        self.files = []
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="TIFF and PDF Metadata Extractor", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File type selection
        file_type_frame = ttk.LabelFrame(main_frame, text="File Type", padding="5")
        file_type_frame.grid(row=1, column=0, columnspan=3, pady=(0, 10), sticky=(tk.W, tk.E))
        
        self.file_type = tk.StringVar(value="tiff")
        tiff_radio = ttk.Radiobutton(file_type_frame, text="TIFF Files", 
                                    variable=self.file_type, value="tiff")
        tiff_radio.grid(row=0, column=0, padx=10)
        
        pdf_radio = ttk.Radiobutton(file_type_frame, text="PDF Files", 
                                   variable=self.file_type, value="pdf")
        pdf_radio.grid(row=0, column=1, padx=10)
        
        # Select files button
        select_btn = ttk.Button(main_frame, text="Select Files", 
                               command=self.select_files)
        select_btn.grid(row=2, column=0, pady=(0, 10), sticky=tk.W)
        
        # Extract button
        extract_btn = ttk.Button(main_frame, text="Extract Metadata", 
                                command=self.extract_metadata)
        extract_btn.grid(row=2, column=1, pady=(0, 10), sticky=tk.W)
        
        # Export button
        export_btn = ttk.Button(main_frame, text="Export to Excel", 
                               command=self.export_to_excel)
        export_btn.grid(row=2, column=2, pady=(0, 10), sticky=tk.E)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Treeview for results
        if self.file_type.get() == "tiff":
            columns = ("Filename", "Format", "Extension", "DPI", "Compression", "Color Depth", "Filename Valid")
        else:
            columns = ("Filename", "Page", "Type", "Color Depth", "DPI", "Compression", "Filename Valid")
            
        self.tree = ttk.Treeview(main_frame, columns=columns, show="headings")
        
        # Set column headings
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        
        # Adjust specific column widths
        self.tree.column("Filename", width=200)
        self.tree.column("Color Depth", width=120)
        self.tree.column("Filename Valid", width=100)
        if "Page" in columns:
            self.tree.column("Page", width=50)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=4, column=3, sticky=(tk.N, tk.S))
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready")
        self.status_label.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E))
        
        # Data storage
        self.metadata_list = []
        
        # Update treeview when file type changes
        self.file_type.trace('w', self.update_columns)
    
    def update_columns(self, *args):
        # Clear the treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Update columns based on file type
        if self.file_type.get() == "tiff":
            columns = ("Filename", "Format", "Extension", "DPI", "Compression", "Color Depth", "Filename Valid")
        else:
            columns = ("Filename", "Page", "Type", "Color Depth", "DPI", "Compression", "Filename Valid")
        
        # Reconfigure treeview
        self.tree.config(columns=columns)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        
        # Adjust specific column widths
        self.tree.column("Filename", width=200)
        self.tree.column("Color Depth", width=120)
        self.tree.column("Filename Valid", width=100)
        if "Page" in columns:
            self.tree.column("Page", width=50)
    
    def select_files(self):
        if self.file_type.get() == "tiff":
            filetypes = (
                ('TIFF files', '*.tif *.tiff'),
                ('All files', '*.*')
            )
        else:
            filetypes = (
                ('PDF files', '*.pdf'),
                ('All files', '*.*')
            )
        
        filenames = filedialog.askopenfilenames(
            title='Open files',
            initialdir='/',
            filetypes=filetypes
        )
        
        if filenames:
            self.files = list(filenames)
            self.status_label.config(text=f"Selected {len(self.files)} files")
            
            # Clear previous results
            for item in self.tree.get_children():
                self.tree.delete(item)
                
            self.metadata_list = []
    
    def extract_metadata(self):
        if not self.files:
            messagebox.showwarning("Warning", "Please select files first")
            return
        
        self.progress.start()
        self.status_label.config(text="Extracting metadata...")
        self.root.update()
        
        # Clear previous results
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        self.metadata_list = []
        
        if self.file_type.get() == "tiff":
            self.extract_tiff_metadata()
        else:
            self.extract_pdf_metadata()
        
        self.progress.stop()
        self.status_label.config(text=f"Extracted metadata from {len(self.metadata_list)} items")
    
    def extract_tiff_metadata(self):
        for file_path in self.files:
            try:
                metadata = self.get_tiff_metadata(file_path)
                self.metadata_list.append(metadata)
                
                # Add to treeview
                self.tree.insert("", "end", values=(
                    metadata["Filename"],
                    metadata["Format"],
                    metadata["Extension"],
                    metadata["DPI"],
                    metadata["Compression"],
                    metadata["Color Depth"],
                    metadata["Filename Valid"]
                ))
                
            except Exception as e:
                self.tree.insert("", "end", values=(
                    os.path.basename(file_path),
                    "Error",
                    "Error",
                    "Error",
                    "Error",
                    "Error",
                    "Error"
                ))
                print(f"Error processing {file_path}: {str(e)}")
    
    def extract_pdf_metadata(self):
        for file_path in self.files:
            try:
                metadata_list = self.get_pdf_metadata(file_path)
                for metadata in metadata_list:
                    self.metadata_list.append(metadata)
                    
                    # Add to treeview
                    self.tree.insert("", "end", values=(
                        metadata["Filename"],
                        metadata["Page"],
                        metadata["Type"],
                        metadata["Color Depth"],
                        metadata["DPI"],
                        metadata["Compression"],
                        metadata["Filename Valid"]
                    ))
                
            except Exception as e:
                self.tree.insert("", "end", values=(
                    os.path.basename(file_path),
                    "Error",
                    "Error",
                    "Error",
                    "Error",
                    "Error",
                    "Error"
                ))
                print(f"Error processing {file_path}: {str(e)}")
    
    def get_tiff_metadata(self, file_path):
        filename = os.path.basename(file_path)
        
        # Check filename convention
        filename_valid = self.check_tiff_filename_convention(filename)
        
        with Image.open(file_path) as img:
            # Get basic info
            img_format = img.format or "Unknown"
            extension = os.path.splitext(file_path)[1].lower()
            
            # Get DPI - extract exact values
            dpi_x, dpi_y = img.info.get('dpi', (0, 0))
            if dpi_x and dpi_y:
                dpi = f"{dpi_x} x {dpi_y}"
            else:
                # Try to get resolution from EXIF data
                dpi = self.get_tiff_resolution(img)
            
            # Get compression - extract exact value
            compression = self.get_tiff_compression(img)
            
            # Get color depth - extract exact value
            color_depth = self.get_tiff_color_depth(img)
        
        return {
            "File Type": "TIFF",
            "Filename": filename,
            "Format": img_format,
            "Extension": extension,
            "DPI": dpi,
            "Compression": compression,
            "Color Depth": color_depth,
            "Filename Valid": filename_valid,
            "Full Path": file_path
        }
    
    def get_tiff_resolution(self, img):
        """Extract exact resolution from TIFF metadata"""
        try:
            # Try to get resolution from EXIF data
            if hasattr(img, '_getexif') and img._getexif():
                exif_data = img._getexif()
                # XResolution tag (282)
                x_res = exif_data.get(282, (1, 1)) if exif_data else (1, 1)
                # YResolution tag (283)
                y_res = exif_data.get(283, (1, 1)) if exif_data else (1, 1)
                
                if isinstance(x_res, tuple) and isinstance(y_res, tuple):
                    x_dpi = round(x_res[0] / x_res[1]) if x_res[1] != 0 else 0
                    y_dpi = round(y_res[0] / y_res[1]) if y_res[1] != 0 else 0
                    if x_dpi and y_dpi:
                        return f"{x_dpi} x {y_dpi}"
            
            # Try to get resolution from tag data
            if hasattr(img, 'tag'):
                tags = img.tag_v2
                # XResolution tag (282)
                if 282 in tags:
                    x_res = tags[282]
                    if isinstance(x_res, tuple) and len(x_res) == 2:
                        x_dpi = round(x_res[0] / x_res[1]) if x_res[1] != 0 else 0
                    else:
                        x_dpi = x_res if isinstance(x_res, (int, float)) else 0
                else:
                    x_dpi = 0
                
                # YResolution tag (283)
                if 283 in tags:
                    y_res = tags[283]
                    if isinstance(y_res, tuple) and len(y_res) == 2:
                        y_dpi = round(y_res[0] / y_res[1]) if y_res[1] != 0 else 0
                    else:
                        y_dpi = y_res if isinstance(y_res, (int, float)) else 0
                else:
                    y_dpi = 0
                
                if x_dpi and y_dpi:
                    return f"{x_dpi} x {y_dpi}"
            
            return "Not specified"
        except:
            return "Not specified"
    
    def get_tiff_compression(self, img):
        """Extract exact compression from TIFF metadata"""
        try:
            compression = "Unknown"
            if hasattr(img, 'tag'):
                tags = img.tag_v2
                if 259 in tags:  # Compression tag
                    compression_code = tags[259]
                    compression_names = {
                        1: "Uncompressed",
                        2: "CCITT 1D",
                        3: "Group 3 Fax",
                        4: "Group 4 Fax",
                        5: "LZW",
                        6: "JPEG",
                        7: "PackBits",
                        8: "Deflate",
                        32773: "PackBits",
                        32946: "Deflate"
                    }
                    compression = compression_names.get(compression_code, f"Unknown ({compression_code})")
            return compression
        except:
            return "Unknown"
    
    def get_tiff_color_depth(self, img):
        """Extract exact color depth from TIFF metadata"""
        try:
            # Get bits per sample from metadata
            bits_per_sample = 0
            if hasattr(img, 'tag'):
                tags = img.tag_v2
                if 258 in tags:  # BitsPerSample tag
                    bits_data = tags[258]
                    if isinstance(bits_data, (list, tuple)):
                        bits_per_sample = bits_data[0] if bits_data else 0
                    else:
                        bits_per_sample = bits_data
            
            # Get samples per pixel from metadata
            samples_per_pixel = 1
            if hasattr(img, 'tag'):
                tags = img.tag_v2
                if 277 in tags:  # SamplesPerPixel tag
                    samples_data = tags[277]
                    if isinstance(samples_data, (list, tuple)):
                        samples_per_pixel = samples_data[0] if samples_data else 1
                    else:
                        samples_per_pixel = samples_data
            
            # Calculate total color depth
            total_bits = bits_per_sample * samples_per_pixel
            
            # Get color mode for additional info
            color_mode = img.mode
            
            if total_bits > 0:
                if color_mode == '1':
                    return "1-bit (Bilevel)"
                elif color_mode in ['L', 'P']:
                    return f"{total_bits}-bit Grayscale"
                elif color_mode == 'RGB':
                    return f"{total_bits}-bit Color"
                elif color_mode == 'RGBA':
                    return f"{total_bits}-bit Color with Alpha"
                elif color_mode == 'CMYK':
                    return f"{total_bits}-bit CMYK"
                else:
                    return f"{total_bits}-bit ({color_mode})"
            else:
                # Fallback to mode-based detection
                if color_mode == '1':
                    return "1-bit (Bilevel)"
                elif color_mode in ['L', 'P']:
                    return "8-bit Grayscale"
                elif color_mode == 'RGB':
                    return "24-bit Color"
                elif color_mode == 'RGBA':
                    return "32-bit Color with Alpha"
                elif color_mode == 'CMYK':
                    return "32-bit CMYK"
                else:
                    return f"Unknown ({color_mode})"
                    
        except:
            return "Unknown"
    
    def get_pdf_metadata(self, file_path):
        filename = os.path.basename(file_path)
        
        # Check filename convention
        filename_valid = self.check_pdf_filename_convention(filename)
        
        metadata_list = []
        
        with open(file_path, 'rb') as file:
            pdf_reader = PdfReader(file)
            num_pages = len(pdf_reader.pages)
            
            for page_num in range(num_pages):
                page = pdf_reader.pages[page_num]
                
                # Extract page metadata
                page_type = "PDF"
                color_depth = "Unknown"
                dpi = "Unknown"
                compression = "Unknown"
                
                # Check if page contains images
                if '/XObject' in page['/Resources']:
                    x_object = page['/Resources']['/XObject'].get_object()
                    for obj in x_object:
                        if x_object[obj]['/Subtype'] == '/Image':
                            page_type = "PDF Image"
                            img = x_object[obj]
                            
                            # Extract color depth information
                            color_depth = self.get_pdf_color_depth(img)
                            
                            # Extract DPI information
                            dpi = self.get_pdf_dpi_info(img, page)
                            
                            # Extract compression information
                            compression = self.get_pdf_compression_info(img)
                            
                            break  # Only analyze the first image on the page
                
                metadata_list.append({
                    "File Type": "PDF",
                    "Filename": filename,
                    "Page": page_num + 1,
                    "Type": page_type,
                    "Color Depth": color_depth,
                    "DPI": dpi,
                    "Compression": compression,
                    "Filename Valid": filename_valid,
                    "Full Path": file_path
                })
        
        return metadata_list
    
    def get_pdf_color_depth(self, img):
        """Extract exact color depth information from PDF image"""
        try:
            color_depth = "Unknown"
            
            # Get color space information
            color_space = "Unknown"
            if '/ColorSpace' in img:
                color_space_obj = img['/ColorSpace']
                if isinstance(color_space_obj, PyPDF2.generic.NameObject):
                    color_space = str(color_space_obj)
                elif isinstance(color_space_obj, PyPDF2.generic.ArrayObject) and len(color_space_obj) > 0:
                    color_space = str(color_space_obj[0])
            
            # Get bits per component
            bits_per_component = 0
            if '/BitsPerComponent' in img:
                bits_per_component = img['/BitsPerComponent']
            
            # Get number of color components
            num_components = self.get_pdf_color_components(color_space)
            
            # Calculate total bits (color depth)
            if bits_per_component and num_components:
                total_bits = bits_per_component * num_components
                color_depth = f"{total_bits}-bit"
                
                # Add color space information
                if color_space == '/DeviceRGB':
                    color_depth += " RGB"
                elif color_space == '/DeviceGray':
                    color_depth += " Grayscale"
                elif color_space == '/DeviceCMYK':
                    color_depth += " CMYK"
                elif color_space == '/Indexed':
                    color_depth += " Indexed"
                else:
                    color_depth += f" ({color_space})"
            
            # If we can't calculate total bits, provide component information
            elif bits_per_component:
                color_depth = f"{bits_per_component}-bit/component"
                if color_space == '/DeviceRGB':
                    color_depth += " RGB"
                elif color_space == '/DeviceGray':
                    color_depth += " Grayscale"
                elif color_space == '/DeviceCMYK':
                    color_depth += " CMYK"
                else:
                    color_depth += f" ({color_space})"
            
            # If no bits per component, just return color space
            elif color_space != "Unknown":
                color_depth = color_space.replace('/', '')
            
            return color_depth
        except Exception as e:
            print(f"Error extracting color depth: {str(e)}")
            return "Unknown"
    
    def get_pdf_color_components(self, color_space):
        """Get number of color components based on color space"""
        if color_space == '/DeviceRGB':
            return 3
        elif color_space == '/DeviceGray':
            return 1
        elif color_space == '/DeviceCMYK':
            return 4
        elif color_space == '/Indexed':
            return 1  # Indexed color typically uses 1 component
        else:
            return 0  # Unknown color space
    
    def get_pdf_dpi_info(self, img, page):
        """Extract exact DPI information from PDF image"""
        try:
            dpi = "Unknown"
            
            # Check if image has width and height
            if '/Width' in img and '/Height' in img:
                width = img['/Width']
                height = img['/Height']
                
                # Check if image has explicit dimensions
                if '/BBox' in img:
                    bbox = img['/BBox']
                    if bbox and len(bbox) == 4:
                        img_width_pt = abs(bbox[2] - bbox[0])
                        img_height_pt = abs(bbox[3] - bbox[1])
                        
                        if img_width_pt > 0 and img_height_pt > 0:
                            dpi_x = round(width / (img_width_pt / 72))
                            dpi_y = round(height / (img_height_pt / 72))
                            dpi = f"{dpi_x} x {dpi_y}"
                
                # If no explicit dimensions, try to use page dimensions
                if dpi == "Unknown" and '/MediaBox' in page:
                    media_box = page['/MediaBox']
                    page_width_pt = abs(media_box[2] - media_box[0])
                    page_height_pt = abs(media_box[3] - media_box[1])
                    
                    # Assume image takes up most of the page
                    if page_width_pt > 0 and page_height_pt > 0:
                        dpi_x = round(width / (page_width_pt / 72))
                        dpi_y = round(height / (page_height_pt / 72))
                        dpi = f"{dpi_x} x {dpi_y}"
            
            return dpi
        except:
            return "Unknown"
    
    def get_pdf_compression_info(self, img):
        """Extract exact compression information from PDF image"""
        try:
            compression = "Unknown"
            if '/Filter' in img:
                filters = img['/Filter']
                if isinstance(filters, PyPDF2.generic.ArrayObject):
                    filter_list = [str(f) for f in filters]
                    compression = ", ".join(filter_list)
                else:
                    compression = str(filters)
                
                # Make compression names more readable
                compression = compression.replace('/FlateDecode', 'Flate')
                compression = compression.replace('/DCTDecode', 'JPEG')
                compression = compression.replace('/JPXDecode', 'JPEG2000')
                compression = compression.replace('/CCITTFaxDecode', 'CCITT')
                compression = compression.replace('/LZWDecode', 'LZW')
                compression = compression.replace('/ASCIIHexDecode', 'ASCIIHex')
                compression = compression.replace('/ASCII85Decode', 'ASCII85')
                compression = compression.replace('/RunLengthDecode', 'RunLength')
            
            return compression
        except:
            return "Unknown"
    
    def check_tiff_filename_convention(self, filename):
        # Check if filename matches ISBN13_#####.tif pattern
        pattern = r'^\d{13}_\d{5}\.tif$'
        if re.match(pattern, filename.lower()):
            return "Yes"
        
        # Check if it's close but has different extension
        pattern_close = r'^\d{13}_\d{5}\.'
        if re.match(pattern_close, filename.lower()):
            return "Wrong extension"
        
        return "No"
    
    def check_pdf_filename_convention(self, filename):
        # Check if filename matches ISBN13.pdf pattern
        pattern = r'^\d{13}\.pdf$'
        if re.match(pattern, filename.lower()):
            return "Yes"
        
        # Check if it's close but has different extension
        pattern_close = r'^\d{13}\.'
        if re.match(pattern_close, filename.lower()):
            return "Wrong extension"
        
        return "No"
    
    def export_to_excel(self):
        if not self.metadata_list:
            messagebox.showwarning("Warning", "No metadata to export. Please extract metadata first.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # Create DataFrame
            df = pd.DataFrame(self.metadata_list)
            
            # Reorder columns based on file type
            if self.file_type.get() == "tiff":
                columns = ["File Type", "Filename", "Format", "Extension", "DPI", "Compression", 
                          "Color Depth", "Filename Valid", "Full Path"]
            else:
                columns = ["File Type", "Filename", "Page", "Type", "Color Depth", "DPI", 
                          "Compression", "Filename Valid", "Full Path"]
            
            # Only include columns that exist in the DataFrame
            columns = [col for col in columns if col in df.columns]
            df = df[columns]
            
            # Export to Excel
            df.to_excel(file_path, index=False)
            
            self.status_label.config(text=f"Exported to {os.path.basename(file_path)}")
            messagebox.showinfo("Success", f"Metadata exported to {file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export to Excel: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MetadataExtractor(root)
    root.mainloop()
