from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, 
                               TableStyle, FrameBreak, NextPageTemplate, Image)
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
import pandas as pd


class PDFTranscriptGenerator:
    def __init__(self, excel_file_path="Database Transkrip Nilai.xlsx"):
        """
        Initialize PDF generator with Excel file path
        
        Args:
            excel_file_path (str): Path to Excel file containing data
        """
        self.excel_file_path = excel_file_path
        self.nilai_data = None
        self.data_siswa = None
        self.data_tetap = None
        self.processed_data_tetap = None
        
        # Load and process data
        self._load_data()
        self._process_data()
    
    def _load_data(self):
        """Load data from Excel file"""
        try:
            data = pd.read_excel(
                self.excel_file_path, 
                sheet_name=["Nilai", "Data Siswa", "Data Tetap"]
            )
            self.nilai_data = data['Nilai']
            self.data_siswa = data['Data Siswa']
            self.data_tetap = data['Data Tetap']
        except Exception as e:
            raise Exception(f"Error loading Excel file: {e}")
    
    def _process_data(self):
        """Process and clean the loaded data"""
        # Clean column names
        self.data_siswa.columns = self.data_siswa.columns.str.strip().str.lower()
        self.nilai_data.columns = self.nilai_data.columns.str.strip().str.lower()
        self.data_tetap.columns = self.data_tetap.columns.str.strip().str.lower()
        
        # Fill NaN values
        self.data_siswa = self.data_siswa.fillna("")
        self.nilai_data = self.nilai_data.fillna("")
        self.data_tetap = self.data_tetap.fillna("")
        
        # Process nilai data
        self.nilai_data = self.nilai_data.set_index('nama').round(2)
        self.nilai_data = self.nilai_data.astype(str).map(lambda x: f"{float(x):.2f}")
        
        # Process student data
        self.data_siswa['tanggal lahir'] = self.data_siswa['tanggal lahir'].apply(
            lambda x: self._datetime_dateId(x)
        )
        self.data_siswa['tempat'] = self.data_siswa['tempat'].apply(
            lambda x: x.capitalize()
        )
        
        # Convert to string
        self.data_siswa = self.data_siswa.astype(str)
        self.nilai_data = self.nilai_data.astype(str)
        self.data_tetap = self.data_tetap.astype(str)
        
        # Process data tetap
        self.data_tetap = self.data_tetap.fillna("").set_index("variabel")
        self.processed_data_tetap = self.data_tetap['nilai'].to_dict()
    
    def _datetime_dateId(self, dt):
        """Convert datetime to Indonesian date format"""
        bln_id = [
            "Januari", "Februari", "Maret", "April", "Mei", "Juni",
            "Juli", "Agustus", "September", "Oktober", "November", "Desember"
        ]
        return f"{dt.day} {bln_id[dt.month - 1]} {dt.year}"
    
    def _buat_tabel(self, row):
        """Create table data from row"""
        data = [["No", "Mata Pelajaran", "Nilai"]]
        data += [[i + 1, col.title(), row[col]] for i, col in enumerate(row.index)]
        return data
    
    def generate_all_pdfs(self):
        """Generate PDF for all students"""
        self.data_siswa.apply(
            lambda x: PDFWithHeaderPadding(
                student_data=x,
                nilai_data=self._buat_tabel(self.nilai_data.loc[x.nama]),
                data_tetap=self.processed_data_tetap
            ),
            axis=1
        )
        print(f"Generated {len(self.data_siswa)} PDF files successfully!")
    
    def generate_single_pdf(self, student_name):
        """Generate PDF for a single student"""
        try:
            student_row = self.data_siswa[self.data_siswa['nama'] == student_name].iloc[0]
            nilai_row = self.nilai_data.loc[student_name]
            
            PDFWithHeaderPadding(
                student_data=student_row,
                nilai_data=self._buat_tabel(nilai_row),
                data_tetap=self.processed_data_tetap
            )
            print(f"Generated PDF for {student_name} successfully!")
        except Exception as e:
            print(f"Error generating PDF for {student_name}: {e}")


class PDFWithHeaderPadding:
    def __init__(self, student_data=None, nilai_data=None, data_tetap=None, output_path=None, ttd_kepsek_path=None):
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
            self.satuan_pendidikan = "PKBM"
            self.kepsek = "John doe"
            self.nipKepsek = "12345678987654321"
            self.tanggal = "Taipei, 17 Agustus 1945"
            self.peminatan = "ipa ips apapun"
            self.tanggal_kelulusan = "xxxxxxx"
        else:
            # Gunakan data dari parameter
            self.kepsek = data_tetap.get('kepsek', 'John doe')
            self.satuan_pendidikan = data_tetap.get('Satuan Pendidikan', 'PKBM')
            self.nipKepsek = data_tetap.get('nip', '12345678987654321')
            self.tanggal = data_tetap.get('Tanggal ttd', 'Taipei, 17 Agustus 1945')
            self.peminatan = data_tetap.get('Peminatan', 'ipa ips apapun')
            self.tanggal_kelulusan = data_tetap.get('Tanggal Kelulusan', 'xxxxxxx')
        
        # DATA SISWA - Set default values atau gunakan parameter
        if student_data is None:
            # Data default
            self.nama = "nama"
            self.tempat_lahir = "ttl"
            self.nisn = "nisn"
            self.no_ijazah = "no ikasah"
            self.no_transkrip = "transk"
            self.no_peserta_ujian = "npsnandanpds"
        else:
            # Format: [nama, tempat_lahir, nisn, no_ijazah, tanggal_kelulusan, peminatan]
            self.nama = student_data.get('nama', 'nama')
            self.tempat_lahir = f"{student_data.get('tempat', 'ttl')}, {student_data.get('tanggal lahir', '')}"
            self.nisn = student_data.get('nisn', 'nisn')
            self.no_ijazah = student_data.get('no ijazah', 'no ikasah')
            self.no_transkrip = student_data.get('no transkrip', 'transk')
            self.no_peserta_ujian = student_data.get('no peserta ujian', 'npsnandanpds')

        # NILAI DATA - Set default values atau gunakan parameter
        if nilai_data is None:
            self.nilai = [
                ["No", "Mata Pelajaran", "Nilai"],
                ["1", "Pendidikan Agama Islam", ""],
                ["2", "Pendidikan Pancasila dan Kewarganegaraan", ""],
                ["3", "Bahasa Indonesia", ""],
                ["4", "Matematika", ""],
                ["5", "Bahasa Inggris", ""],
                ["6", "Sejarah Indonesia", ""],
                ["7", "Sosiologi", ""],
                ["8", "Geografi", ""],
                ["9", "Ekonomi", ""],
                ["10", "Sejarah", ""],
                ["11", "Keterampilan Wajib", ""],
                ["12", "Keterampilan Pilihan", ""],
                ["13", "Pemberdayaan", ""]
            ]
        else:
            self.nilai = nilai_data

        self.styles = getSampleStyleSheet()
        self.story = []
        self._setup_fonts()
        self._setup_style()
        
        # Use output_path if provided, otherwise use default
        if output_path:
            self.filename = output_path
            # Create directory if needed
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
        else:
            self.filename = f"transkrip/transkript_{self.nama}.pdf"
        
        self.build_pdf()

    def _setup_fonts(self):
        """Setup font configuration"""
        # Get the project root directory (2 levels up from this file)
        project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        self.fontpath = os.path.join(project_root, "assets", "fonts")
        
        try:
            pdfmetrics.registerFont(TTFont('TimesRoman', os.path.join(self.fontpath, 'times.ttf')))
            pdfmetrics.registerFont(TTFont('TimesRomanBold', os.path.join(self.fontpath, 'timesbd.ttf')))
            pdfmetrics.registerFont(TTFont('TimesRomanBoldItalic', os.path.join(self.fontpath, 'timesbdit.ttf')))
            pdfmetrics.registerFont(TTFont('TimesRomanItalic', os.path.join(self.fontpath, 'timesit.ttf')))

            self.times = 'TimesRoman'
            self.timesBold = 'TimesRomanBold'
            self.timesBoldItalic = 'TimesRomanBoldItalic'
            self.timesItalic = 'TimesRomanItalic'
        except:
            print("Warning: Font tidak ditemukan, menggunakan font default")
            self.times = 'Times-Roman'
            self.timesBold = 'Times-Roman'
            self.timesBoldItalic = 'Times-Roman'
            self.timesItalic = 'Times-Roman'
    
    def _setup_style(self):
        """Setup paragraph styles"""
        self.styles = getSampleStyleSheet()
        
        self.center_ttd = ParagraphStyle(
            'center_ttd',
            parent=self.styles['Normal'],
            fontName=self.times,
            fontSize=12,
            alignment=1,
            leading=15
        )
        
        self.center_ttd_bold = ParagraphStyle(
            'center_ttd_bold',
            parent=self.styles['Normal'],
            fontName=self.timesBold,
            fontSize=12,
            alignment=1,
            leading=15
        )
        
        # Footer style
        self.headerItalic = ParagraphStyle(
            'CustomFooter',
            parent=self.styles['Normal'],
            fontName=self.timesBoldItalic,
            fontSize=8,
            leading=15,
            alignment=1
        )
        
        # Cover bold style
        self.headerBold = ParagraphStyle(
            'cover',
            parent=self.styles['Normal'],
            fontName=self.timesBold,
            fontSize=13,
            leading=20,
            alignment=1
        )
        
        self.title = ParagraphStyle(
            'title',
            parent=self.styles['Normal'],
            fontName=self.timesBold,
            fontSize=14,
            leading=18,
            alignment=1
        )
        
        self.noTitle = ParagraphStyle(
            'notitle',
            parent=self.styles['Normal'],
            fontName=self.times,
            fontSize=11,
            leading=15,
            alignment=1
        )
        
    def add_spacer(self, height_cm=0.5):
        """Menambah spacer dengan tinggi tertentu dalam cm"""
        self.story.append(Spacer(1, height_cm * cm))
    
    def bold_underline(self, text):
        """Format text with bold and underline"""
        return f"<b><u>{text}</u></b>"
    
    def ttdKananKepsek(self):
        """Generate signature section for principal"""
        title = "Kepala,"
        nama = Paragraph(self.bold_underline(self.kepsek), self.center_ttd_bold)
        nip = Paragraph("NIP. " + self.nipKepsek, self.center_ttd)
                
        ttd_kepsek_style = [
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, -1), self.times),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
        ]       
        
        # Build signature content with optional image
        if self.ttd_kepsek_img:
            keterangan = Paragraph(
                f"{self.tanggal}<br/>{title}", 
                self.center_ttd
            )
            text = [keterangan, self.ttd_kepsek_img, nama, nip]
        else:
            keterangan = Paragraph(
                f"{self.tanggal}<br/>{title}<br/><br/><br/><br/><br/>", 
                self.center_ttd
            )
            text = [keterangan, nama, nip] 
        
        ttd_kepsek = [["", "", text,""]]
        ttd_kepsek_table = Table(
            ttd_kepsek, 
            colWidths=["*", 4*cm, 9*cm,1.5*cm], 
            rowHeights=[4*cm]
        )
        ttd_kepsek_table.setStyle(TableStyle(ttd_kepsek_style))        
        
        self.story.append(ttd_kepsek_table)
            
    def build_pdf(self):
        """Build the PDF document"""
        # Create directory if it doesn't exist
        os.makedirs("transkrip", exist_ok=True)
        
        doc = BaseDocTemplate(self.filename, pagesize=A4)

        # FRAME HEADER (tinggi 4 cm, mulai dari 1 cm dari atas)
        header_frame = Frame(
            x1=2*cm,
            y1=A4[1] - 6*cm,
            width=A4[0] - 4*cm,
            height=5*cm,
            leftPadding=0, bottomPadding=0, rightPadding=0, topPadding=0,
            id='header_frame'
        )

        # FRAME ISI (mulai 1 cm di bawah header, jadi dari 6 cm dari atas)
        content_frame = Frame(
            x1=1*cm,
            y1=0.5*cm,
            width=A4[0] - 2*cm,
            height=A4[1] - 5.8*cm,
            leftPadding=0, bottomPadding=0, rightPadding=0, topPadding=0,
            id='content_frame'
        )

        def header_story(canvas, doc):
            pass

        template = PageTemplate(
            id='main_template',
            frames=[header_frame, content_frame],
            onPage=header_story
        )

        doc.addPageTemplates([template])

        # Load logo
        try:
            project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            logo_path = os.path.join(project_root, "assets", "logo.png")
            img = Image(logo_path, width=4*cm, height=4*cm)
            img.hAlign = 'CENTER'
        except Exception as e:
            print(f"Error loading logo: {e}")
            img = None
            
        # === HEADER CONTENT ===
        text1 = Paragraph(
            """Yayasan Pendidikan Perhimpunan Pelajar Indonesia Taiwan<br/>
            PUSAT KEGIATAN BELAJAR MASYARAKAT (PKBM) /<br/> COMMUNITY LEARNING CENTER (CLC)-PPI TAIWAN<br/>""", 
            self.headerBold
        )
        
        text2 = Paragraph(
            """KDEI 6F, No. 550, Rui Guang Road, Neihu District, Taipei, 114, Taiwan, ROC<br/>
            Website: www.pkbmppitaiwan.sch.id; Telp: - Email: ketua@pkbmppitaiwan.sch.id<br/>
            Nomor Induk Lembaga (NILEM): P9908880""", 
            self.headerItalic
        )
        
        if img:
            header_data = [[img, [text1, text2]]]
            header_table = Table(header_data, colWidths=[4*cm, "*"])
        else:
            header_data = [[[text1, text2]]]
            header_table = Table(header_data, colWidths=["*"])
            
        header_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 0), (0, 0), 'LEFT'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('LEFTPADDING', (1, 0), (1, 0), 10),
            ('RIGHTPADDING', (1, 0), (1, 0), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTSIZE', (0, 0), (-1, -1), 12),
        ]))

        self.story.append(header_table)
        self.story.append(FrameBreak())

        # === ISI KONTEN ===
        self.story.append(Paragraph(
            "TRANSKRIP NILAI<br/>PENDIDIKAN KESETARAAN PAKET C", 
            self.title
        ))
        self.story.append(Paragraph(f"Nomor: {self.no_transkrip}", self.noTitle))
        self.add_spacer()
        
        # Student data table
        data = [
            ["Satuan Pendidikan", ":", self.satuan_pendidikan],
            ["Nomor Pokok Sekolah Nasional", ":", "P9908880"],
            ["Nama Lengkap", ":", self.nama],
            ["Tempat, Tanggal Lahir", ":", self.tempat_lahir],
            ["Nomor Induk Siswa Nasional", ":", self.nisn],
            ["Nomor Ijazah", ":", self.no_ijazah],
            ["Tanggal Kelulusan", ":", self.tanggal_kelulusan],
            ["Peminatan", ":", self.peminatan]
        ]
        
        table1 = Table(data, colWidths=[6.2*cm, 0.2*cm, 5.5*cm])
        table1.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
        ]))

        # Grades table
        self.story.append(table1)
        self.add_spacer()
        
        table2 = Table(self.nilai, colWidths=[30, 230, 70])
        table2.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), self.times),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('FONTNAME', (0, 0), (-1, 0), self.timesBold),
            ('FONTSIZE', (0, 0), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('ALIGN', (0, 0), (0, -1), 'CENTER'),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('ALIGN', (-1, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        self.story.append(table2)
        self.add_spacer(1)
        
        # Signature
        self.ttdKananKepsek()
        doc.build(self.story)


def generate_all_transcripts(excel_file_path, output_folder, ttd_kepsek_path=None, progress_callback=None):
    """
    Wrapper function for GUI compatibility.
    Generate Transkrip Nilai PDFs for all students from Excel file.
    
    Args:
        excel_file_path: Path to Excel file containing transcript data
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
        
        # Load and generate
        generator = PDFTranscriptGenerator(excel_file_path)
        
        # Get student count for progress
        total_students = len(generator.data_siswa)
        
        # Generate for each student
        count = 0
        for idx, row in generator.data_siswa.iterrows():
            nama = row['nama']
            try:
                pdf_path = os.path.join(output_folder, f"Transkrip_{nama}.pdf")
                # Generate single PDF
                PDFWithHeaderPadding(
                    student_data=row,
                    nilai_data=generator._buat_tabel(generator.nilai_data.loc[nama]),
                    data_tetap=generator.processed_data_tetap,
                    output_path=pdf_path,
                    ttd_kepsek_path=ttd_kepsek_path
                )
                generated_files.append(pdf_path)
            except Exception as e:
                print(f"Error generating Transkrip for {nama}: {e}")
            
            count += 1
            if progress_callback:
                progress_callback(count, total_students, nama)
        
        print(f"Generated {len(generated_files)} Transkrip files")
        
    except Exception as e:
        print(f"Error in generate_all_transcripts: {str(e)}")
        import traceback
        traceback.print_exc()
    
    return generated_files


if __name__ == "__main__":
    # Usage example:
    # 1. Generate all PDFs
    generator = PDFTranscriptGenerator("Database Transkrip Nilai.xlsx")
    generator.generate_all_pdfs()
    
    # 2. Generate single PDF
    # generator.generate_single_pdf("Aan Antisah")