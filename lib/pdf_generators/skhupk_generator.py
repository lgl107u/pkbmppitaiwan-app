from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, Image
from reportlab.lib.enums import TA_LEFT, TA_CENTER
import pandas as pd
import os

class PDFGenerator:
    def __init__(self, student_data=None, nilai_data=None, data_tetap=None, ttd_kepsek_path=None):
        # Page layout settings
        self.page_width, self.page_height = A4
        self.border_margin = 1 * cm
        self.inner_margin = 0.9 * cm
        self.inner_margin_top_bottom = 0.5 * cm
        self.content_margin = self.border_margin + self.inner_margin
        self.top_bottom = self.border_margin + self.inner_margin_top_bottom
        self.content_width = self.page_width - 2 * self.content_margin
        
        # Get project root directory
        project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        
        # Logo settings
        self.logo_filename = os.path.join(project_root, "assets", "logo.png")
        self.logo_width = 3 * cm
        self.logo_height = 3 * cm
        
        # Signature image
        self.ttd_kepsek_path = ttd_kepsek_path
        self.ttd_kepsek_img = None
        if ttd_kepsek_path and os.path.exists(ttd_kepsek_path):
            try:
                self.ttd_kepsek_img = Image(ttd_kepsek_path, width=3.5*cm, height=2*cm)
            except Exception as e:
                print(f"Error loading signature: {e}")
        
        # DATA TETAP - Set default values atau gunakan parameter
        if data_tetap is None:
            # Data tetap default
            self.default_tahun = "TEMPLATE"
            self.penyelenggara_upk = "TEMPLATE"
            self.satuan_pendidikan_asal = "TEMPLATE"
            self.signature_date_placeholder = "TEMPLATE"
            self.signature_title = "Kepala PKBM"
            self.signature_name = "TEMPLATE"
            self.signature_nip = "TEMPLATE"
        else:
            # Gunakan data dari parameter
            self.default_tahun = data_tetap.get('tahun')
            self.penyelenggara_upk = data_tetap.get('Penyelenggara UPK')
            self.satuan_pendidikan_asal = data_tetap.get('Satuan Pendidikan Asal')
            self.signature_date_placeholder = data_tetap.get('tanggal')
            self.signature_title = "Kepala PKBM"
            self.signature_name = data_tetap.get('kepsek')
            self.signature_nip = "NIP. "+ str(data_tetap.get('nip'))
        
        self.photo_placeholder = "Foto Siswa \n\n\n3x4"
        
        # DATA BERUBAH - Set default values atau gunakan parameter
        if student_data is None:
            # Data default
            self.nama = "TEMPLATE"
            self.tempat = "TEMPLATE"
            self.tanggal_lahir = "TEMPLATE"
            self.nisn = "TEMPLATE"
            self.no_peserta = "TEMPLATE"
            self.document_number = "TEMPLATE"
        else:
            # Gunakan data dari parameter
            # Format: [nama, tempat, tanggal_lahir, nisn, no_peserta, document_number]
            self.nama = student_data[0] if len(student_data) > 0 else "TEMPLATE"
            self.tempat = student_data[1] if len(student_data) > 1 else "TEMPLATE"
            self.tanggal_lahir = student_data[2] if len(student_data) > 2 else "TEMPLATE"
            self.nisn = student_data[3] if len(student_data) > 3 else "TEMPLATE"
            self.no_peserta = student_data[4] if len(student_data) > 4 else "TEMPLATE"
            self.document_number = student_data[5] if len(student_data) > 5 else "TEMPLATE"
        
        self.document_title = f"SURAT KETERANGAN<br/>NO.{self.document_number}"
        self.header_intro = "Kepala Pusat Kegiatan Belajar Masyarakat (PKBM) PPI Taiwan, menerangkan bahwa:"
        
        # Exam section data
        self.exam_section_title = "TELAH MENGIKUTI"
        self.exam_program = "Ujian Pendidikan Kesetaraan (UPK) Program Paket C, Peminatan Ilmu-Ilmu Sosial (IIS)"
        
        # Default grades data atau gunakan parameter
        if nilai_data is None:
            self.default_grades_data = [
                ['No', 'Mata Pelajaran', 'Nilai'],
                ['1', 'TEMPLATE', 89],
                ['2', 'TEMPLATE', 88],
                ['3', 'TEMPLATE', 76],
                ['4', 'TEMPLATE', 76],
                ['5', 'TEMPLATE', 98],
                ['5', 'TEMPLATE', 98],
                ['5', 'TEMPLATE', 98],
                ['5', 'TEMPLATE', 98],
                ['5', 'TEMPLATE', 98],
                ['5', 'TEMPLATE', 98],
                ['5', 'TEMPLATE', 98],
                ['5', 'TEMPLATE', 98],
                ['5', 'TEMPLATE', 98],
            ]
        else:
            self.default_grades_data = nilai_data
        
        
        # Table column widths
        self.data_table_col1_width = 0.3  # 30% of content width
        self.data_table_col2_width = 0.01  # 1% of content width
        self.nilai_table_col_widths = [1 * cm, 8.5 * cm, 1.5 * cm]
        self.ttd_table_col_widths = ["*", "*", 3*cm, 7.5*cm]
        self.ttd_table_row_height = 4 * cm
        
        # Initialize styles and story
        self.styles = getSampleStyleSheet()
        self.story = []
        self._setup_styles()
        
        # Store output_folder if provided, don't auto-generate
        self.output_folder = None
    
    def _setup_styles(self):
        """Setup custom styles untuk berbagai elemen"""
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=18,
            alignment=TA_CENTER,
            textColor=colors.black,
            fontName='Helvetica-Bold'
        )
        
        self.header_style = ParagraphStyle(
            'Header',
            parent=self.styles['Normal'],
            fontSize=12,
            leading=16,
            alignment=TA_LEFT,
            textColor=colors.black,
            fontName='Helvetica'
        )

        self.text_style = ParagraphStyle(
            'Text',
            parent=self.styles['Normal'],
            fontSize=15,
            leading=20,
            spaceAfter=15,
            alignment=TA_LEFT,
            textColor=colors.black,
            fontName='Helvetica'
        )
        
        self.grey_center = ParagraphStyle(
            'GreyCenter',
            parent=self.styles['Normal'],
            fontSize=15,
            leading=20,
            alignment=TA_CENTER,
            textColor=colors.black,
            fontName='Helvetica'
        )
        
        self.padding = ParagraphStyle(
            'Padding',
            parent=self.styles['Normal'],
            alignment=TA_CENTER,
            fontSize=10,
            leading=14,
            leftIndent=4
        )
    
    def draw_border(self, canvas, doc):
        """Menggambar border pada halaman"""
        canvas.saveState()
        canvas.setLineWidth(1)
        canvas.rect(
            self.border_margin,
            self.border_margin,
            self.page_width - 2 * self.border_margin,
            self.page_height - 2 * self.border_margin
        )
        canvas.restoreState()
    
    def add_spacer(self, height_cm=0.5):
        """Menambah spacer dengan tinggi tertentu dalam cm"""
        self.story.append(Spacer(1, height_cm * cm))
    
    def add_title(self, title_text=None):
        """Menambah judul dokumen"""
        if title_text is None:
            title_text = self.document_title
        title = Paragraph(title_text, self.title_style)
        self.story.append(title)
    
    def add_header_text(self, text=None):
        """Menambah teks header"""
        if text is None:
            text = self.header_intro
        header = Paragraph(text, self.header_style)
        self.story.append(header)
    
    def create_data_table(self, data=None):
        """Membuat tabel data siswa dengan format yang rapi"""            
        table_data = [
            ['Nama', ':', self.nama],
            ['Tempat/Tgl. Lahir', ':', f"{self.tempat}, {self.tanggal_lahir}"],
            ['NIPD/NISN', ':', str(self.nisn)],
            ['No. Peserta Ujian', ':', str(self.no_peserta)],
            ['Penyelenggara UPK', ':', self.penyelenggara_upk],
            ['Satuan Pendidikan Asal', ':', self.satuan_pendidikan_asal]
        ]
        
        col1_width = self.content_width * self.data_table_col1_width
        col2_width = self.content_width * self.data_table_col2_width
        
        table = Table(table_data, colWidths=[col1_width, col2_width, "*"])
        table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTNAME', (1, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 12),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),
            ('ALIGN', (2, 0), (2, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('LEFTPADDING', (0, 0), (0, -1), 0),
            ('LEFTPADDING', (2, 0), (2, -1), 8),
        ]))
        
        return table
    
    def create_nilai_table(self, data=None):
        """Membuat tabel nilai dengan grid"""
        if data is None:
            data = self.default_grades_data
            
        table = Table(data, colWidths=self.nilai_table_col_widths)
        
        table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("ALIGN", (1, 1), (1, -1), "LEFT"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("GRID", (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        return table
    
    def create_ttd_table(self, data=None):
        """Membuat tabel tanda tangan"""
        if data is None:
            # Build signature cell with optional image
            if self.ttd_kepsek_img:
                signature_content = [
                    Paragraph(f"{self.signature_date_placeholder}<br/>{self.signature_title}", self.padding),
                    self.ttd_kepsek_img,
                    Paragraph(f"<b>{self.signature_name}</b><br/>{self.signature_nip}", self.padding)
                ]
                # Create inner table for signature with image
                sig_table = Table([[sig] for sig in signature_content], colWidths=[7*cm])
                sig_table.setStyle(TableStyle([
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ]))
                data = [[" ", " ", self.photo_placeholder, sig_table]]
            else:
                signature_text = f"{self.signature_date_placeholder}<br/>{self.signature_title}<br/><br/><br/><br/><br/><b>{self.signature_name}</b><br/>{self.signature_nip}"
                data = [[
                    " ", " ", self.photo_placeholder,
                    Paragraph(signature_text, self.padding)
                ]]
        
        table = Table(data, colWidths=self.ttd_table_col_widths, rowHeights=[self.ttd_table_row_height])
        
        table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("GRID", (2, 0), (2, 0), 1, colors.black),
            ("LEFTPADDING", (2, 0), (2, 0), 4),
        ]))
        
        return table
    
    def generate_pdf(self, filename="document.pdf", data=None, tahun=None):
        """Generate PDF dengan data yang diberikan"""
        # Set default values jika data tidak disediakan
        if tahun is None:
            tahun = self.default_tahun
        
        # Reset story untuk generate baru
        self.story = []
        
        # Buat dokumen PDF
        doc = SimpleDocTemplate(
            filename,
            pagesize=A4,
            leftMargin=self.content_margin,
            rightMargin=self.content_margin,
            topMargin=self.top_bottom,
            bottomMargin=self.top_bottom,
        )
        
        # Tambahkan konten secara berurutan
        self._add_header_section()
        self._add_student_data_section(data)
        self._add_exam_info_section(tahun)
        self._add_grades_section()
        self._add_signature_section()
        
        # Build PDF dengan border
        doc.build(self.story, onFirstPage=self.draw_border, onLaterPages=self.draw_border)
        print(f"PDF berhasil dibuat: {filename}")
    
    def add_logo(self):
        """Menambahkan logo ke dokumen"""
        try:
            img = Image(self.logo_filename, width=self.logo_width, height=self.logo_height)
            img.hAlign = 'CENTER'
            self.story.append(img)
        except Exception as e:
            print(f"Error loading logo: {e}")
    
    def _add_header_section(self):
        """Menambahkan bagian header (logo dan judul)"""
        self.add_logo()
        self.add_spacer(0.3)
        self.add_title()
        self.add_spacer(0.3)
        self.add_header_text()
    
    def _add_student_data_section(self, data):
        """Menambahkan bagian data siswa"""
        data_table = self.create_data_table(data)
        self.story.append(data_table)
    
    def _add_exam_info_section(self, tahun):
        """Menambahkan bagian informasi ujian"""
        self.add_spacer()
        self.story.append(Paragraph(f'<span backColor="#cecece">&nbsp;{self.exam_section_title}&nbsp;</span>', self.grey_center))
        self.add_spacer()
        
        exam_info_text = f"{self.exam_program}, Tahun Pelajaran {tahun}, sesuai dengan peraturan perundang-undangan dengan nilai sebagai berikut:"
        self.story.append(Paragraph(exam_info_text, self.header_style))
        self.add_spacer(0.3)
    
    def _add_grades_section(self, data=None):
        """Menambahkan bagian tabel nilai"""
        grades_table = self.create_nilai_table(data)
        self.story.append(grades_table)
    
    def _add_signature_section(self):
        """Menambahkan bagian tanda tangan"""
        self.add_spacer()
        signature_table = self.create_ttd_table()
        self.story.append(signature_table)

def datetime_dateId(dt):
    bln_id = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    return f"{dt.day} {bln_id[dt.month - 1]} {dt.year}"

def tabel_nilai(row):
    header = ["No" ,"Mata Pelajaran", "Nilai" ]
    kolom_mapel = row.index[1:-1]   # Lewatkan kolom pertama dan terakhir
    data = [[i + 1, col, int(row[col])] for i, col in enumerate(kolom_mapel)]
    # Baris terakhir tanpa nomor
    data.append(["", row.index[-1], int(row[row.index[-1]]) ])
    return [header] + data

if __name__ == "__main__":    
    df = pd.read_excel("./Database  SKHUPK 2025.xlsx", sheet_name=None)
    
    data_siswa = df['DATA SISWA']

    df['NILAI']["Rata-rata"] = df["NILAI"].iloc[:,1:].mean(axis=1).apply(lambda x: int(round(x,0)))
    df["NILAI"].iloc[:,1:] = df["NILAI"].iloc[:,1:].round().astype(int)
    
    nilai_df = df['NILAI'].set_index('Nama')
    
    # df['NILAI'].set_index('Nama').loc["Aan Antisah"]
    df['DATA TETAP']['variabel'] = df['DATA TETAP']['variabel'].str.strip()
    df['DATA TETAP'] = df['DATA TETAP'].fillna("")
    
    data_tetap = df['DATA TETAP'].set_index('variabel')['nilai'].to_dict()
        
    data_siswa.apply(lambda x: PDFGenerator( [ x.Nama, x.Tempat.capitalize(), datetime_dateId(x['Tanggal Lahir']), x.NISN, x['No peserta Ujian'],  x['Nomor SKHUPK']] , tabel_nilai(nilai_df.loc[ x.Nama ]) , data_tetap), axis=1)


def generate_all_skhupk(excel_file_path, output_folder, ttd_kepsek_path=None, progress_callback=None):
    """
    Wrapper function for GUI compatibility.
    Generate SKHUPK PDFs for all students from Excel file.
    
    Args:
        excel_file_path: Path to Excel file containing SKHUPK data
        output_folder: Output folder for generated PDFs
        ttd_kepsek_path: Path to kepsek signature image (optional)
        progress_callback: Optional callback(current, total, student_name)
    
    Returns:
        List of generated file paths
    """
    import warnings
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
    
    generated_files = []
    
    try:
        # Create output folder
        os.makedirs(output_folder, exist_ok=True)
        
        # Load Excel data
        df = pd.read_excel(excel_file_path, sheet_name=None)
        
        # Process data
        data_siswa = df['DATA SISWA']
        total_students = len(data_siswa)
        
        # Process nilai
        df['NILAI']["Rata-rata"] = df["NILAI"].iloc[:,1:].mean(axis=1).apply(lambda x: int(round(x,0)))
        df["NILAI"].iloc[:,1:] = df["NILAI"].iloc[:,1:].round().astype(int)
        nilai_df = df['NILAI'].set_index('Nama')
        
        # Process data tetap
        df['DATA TETAP']['variabel'] = df['DATA TETAP']['variabel'].str.strip()
        df['DATA TETAP'] = df['DATA TETAP'].fillna("")
        data_tetap = df['DATA TETAP'].set_index('variabel')['nilai'].to_dict()
        
        # Generate for each student
        count = 0
        for idx, row in data_siswa.iterrows():
            nama = row['Nama']
            student_data = [
                nama,
                row['Tempat'].capitalize(),
                datetime_dateId(row['Tanggal Lahir']),
                row['NISN'],
                row['No peserta Ujian'],
                row['Nomor SKHUPK']
            ]
            nilai_data = tabel_nilai(nilai_df.loc[nama])
            
            try:
                pdf_path = os.path.join(output_folder, f"SKHUPK_{nama}.pdf")
                generator = PDFGenerator(student_data, nilai_data, data_tetap, ttd_kepsek_path=ttd_kepsek_path)
                generator.generate_pdf(pdf_path)
                generated_files.append(pdf_path)
            except Exception as e:
                print(f"Error generating SKHUPK for {nama}: {e}")
            
            count += 1
            if progress_callback:
                progress_callback(count, total_students, nama)
        
        print(f"Generated {len(generated_files)} SKHUPK files")
        
    except Exception as e:
        print(f"Error in generate_all_skhupk: {str(e)}")
        import traceback
        traceback.print_exc()
    
    return generated_files
