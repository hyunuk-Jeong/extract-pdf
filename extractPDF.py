import os
import fitz  # PyMuPDF
import requests
import re
from urllib.parse import unquote
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk  # ttk를 사용하여 더 나은 스타일 적용

# 폴더 경로 설정
pdf_folder = ""

# 파일을 저장할 폴더를 각 PDF 파일이 있는 폴더 내에 동적으로 생성
def get_save_folder(pdf_path):
    # PDF 파일이 있는 폴더 내에 'files' 폴더를 생성
    save_folder = os.path.join(os.path.dirname(pdf_path), "files")
    
    # 저장 폴더가 없으면 생성
    os.makedirs(save_folder, exist_ok=True)
    return save_folder

# Windows에서 유효하지 않은 문자를 처리하는 함수
def sanitize_filename(filename):
    # 파일 이름에서 유효하지 않은 문자 제거
    invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename

def get_filename_from_url(url):
    try:
        # URL에 대해 HEAD 요청을 보내서 파일 이름을 확인
        response = requests.head(url, allow_redirects=True, timeout=10)
        # 'Content-Type' 헤더에서 파일인지 웹 페이지인지 확인
        content_type = response.headers.get('Content-Type', '')
        if 'text/html' in content_type:
            print(f"The URL points to a webpage, not a file.\nUrl: {url}")
            return None
        # 'Content-Disposition' 헤더에서 파일 이름 추출
        if 'Content-Disposition' in response.headers:
            content_disposition = response.headers['Content-Disposition']
            
            # 'filename' 또는 'filename*'에서 파일 이름을 추출
            match = re.search(r'filename\*?=(?:"?UTF-8\'\'''?|)(.*)', content_disposition)
            
            # 파일 이름 추출
            if match:
                filename = match.group(1)  # 파일 이름 부분
                filename = unquote(filename)  # URL 디코딩
                filename = sanitize_filename(filename)  # 유효하지 않은 문자 처리
                return filename.strip('"')
        # 없으면 URL에서 추출한 파일 이름을 반환
        return sanitize_filename(url.split('/')[-1])
    except Exception as e:
        print(f"Error extracting filename from URL: {e}")
        return None

# 파일 중복 체크 후 이름 생성
def generate_unique_filename(save_folder, filename):
    base_name, ext = os.path.splitext(filename)
    new_filename = filename
    counter = 1
    
    # 같은 이름의 파일이 존재하면 (1), (2) 식으로 이름을 바꿔서 저장
    while os.path.exists(os.path.join(save_folder, new_filename)):
        new_filename = f"{base_name} ({counter}){ext}"
        counter += 1
    
    return new_filename

# 확장자가 없는 파일을 확인하는 함수
def has_extension(filename):
    # 파일명이 최소 5글자 이상인 경우
    if len(filename) >= 5:
        # 마지막 5글자에 마침표가 없으면 확장자가 없다고 판단
        if '.' not in filename[-5:]:
            return False
        
    return '.' in filename

def download_file(url, local_path, progress_label_download_file):
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()  # 요청이 성공했는지 확인
        
        total_size = int(response.headers.get('content-length', 0))  # 전체 파일 크기 (바이트)
        downloaded_size = 0
        
        with open(local_path, 'wb') as file:
            for data in response.iter_content(chunk_size=8192):
                file.write(data)
                
                downloaded_size += len(data)
                
                # MB 단위로 변환
                downloaded_mb = downloaded_size / (1024 * 1024)  # 다운로드 받은 크기 (MB)
                total_mb = total_size / (1024 * 1024)  # 전체 파일 크기 (MB)
                progress = (downloaded_size / total_size) * 100  # 다운로드 진행률 (%)

                # 진행 상태 업데이트
                progress_label_download_file.config(
                    text=f"링크 다운로드 현황: {downloaded_mb:.2f} MB / {total_mb:.2f} MB ({progress:.2f}%)"
                )
                
                progress_label_download_file.update()
                
        print(f"Downloaded {url} to {local_path}")
    except Exception as e:
        print(f"Failed to download {url}: {e}")

def extract_and_download(pdf_path, total_files, processed_files, download_count, file_count_label, progress_label_pdf, progress_label_download,progress_label_download_file):
    # PDF 파일 열기
    document = fitz.open(pdf_path)
    
    # 파일 저장 폴더 설정
    save_folder = get_save_folder(pdf_path)
    
    # 전체 링크 수 추적
    total_links = 0
    download_count = 0
    
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        
        # 페이지의 하이퍼링크 정보를 가져오기
        links = page.get_links()
        total_links += len(links)
        
        for link in links:
            if 'uri' in link:
                try:
                    link_text = page.get_text("text", clip=link['from'])
                    link_url = link['uri']
                    
                    filename = get_filename_from_url(link_url)
                    
                    # 파일 이름 호출 실패 시 통과
                    if filename == None:
                        download_count += 1
                        # 진행 상태 업데이트 (다운로드 진행)
                        progress_label_download.config(text=f"다운로드 진행: {download_count}/{total_links} 개 링크 다운로드 중")
                        progress_label_download.update()  # GUI 업데이트
                        continue
                    
                    # 파일 이름 중복 체크
                    filename = generate_unique_filename(save_folder, filename)
                    local_path = os.path.join(save_folder, filename)
                    
                    progress_label_download_fileName.config(text=f"파일 이름: {filename}")
                    progress_label_download_fileName.update()  # GUI 업데이트
                    
                    
                    # 확장자가 없는 파일은 다운로드하지 않음
                    if not has_extension(filename):
                        download_count += 1
                        # 진행 상태 업데이트 (다운로드 진행)
                        progress_label_download.config(text=f"다운로드 진행: {download_count}/{total_links} 개 링크 다운로드 중")
                        progress_label_download.update()  # GUI 업데이트
                        continue
                    
                    # 파일 다운로드
                    download_file(link_url, local_path, progress_label_download_file)
                    
                    download_count += 1
                    # 진행 상태 업데이트 (다운로드 진행)
                    progress_label_download.config(text=f"다운로드 진행: {download_count}/{total_links} 개 링크 다운로드 중")
                    progress_label_download.update()  # GUI 업데이트
                except Exception as e:
                    download_count += 1
                    print(f"Error processing {pdf_path}: {e}")
                    continue                
    
    processed_files += 1
    # 진행 중인 파일 업데이트 (PDF 처리 진행)
    progress_label_pdf.config(text=f"진행 중: {processed_files}/{total_files} 개 PDF 파일 처리 중\n현재 PDF: {os.path.basename(pdf_path)}")
    progress_label_pdf.update()

    document.close()
    

# PDF 폴더 내 모든 파일 처리 (files 폴더 제외)
def process_pdfs(file_count_label, progress_label_pdf, progress_label_download,progress_label_download_file,progress_label_download_fileName):
    global pdf_folder
    
    start_button.config(text="진행중...", state=tk.DISABLED)  # 버튼 텍스트 변경 및 비활성화
    
    
    
    if not pdf_folder:
        messagebox.showerror("Error", "PDF 폴더를 선택해 주세요.")
        start_button.config(text="시작", state=tk.NORMAL)  # 버튼 텍스트 변경 및 활성화
        return
    
    # 전체 파일 개수 계산
    total_files = 0
    for root, dirs, files in os.walk(pdf_folder):
        if "files" in dirs:
            dirs.remove("files")
        for file in files:
            if file.lower().endswith('.pdf'):
                total_files += 1
    
    if total_files == 0:
        messagebox.showerror("Error", "PDF 파일이 없습니다.")
        start_button.config(text="시작", state=tk.NORMAL)  # 버튼 텍스트 변경 및 활성화
        return

    # 파일 수 표시
    file_count_label.config(text=f"전체 PDF 파일: {total_files} 개")
    
    # 폴더 내 PDF 파일 처리
    processed_files = 0
    total_links = 0
    for root, dirs, files in os.walk(pdf_folder):
        if "files" in dirs:
            dirs.remove("files")
        
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_path = os.path.join(root, file)
                # PDF 파일 내의 링크 개수 계산
                document = fitz.open(pdf_path)
                for page_num in range(len(document)):
                    page = document.load_page(page_num)
                    links = page.get_links()
                    total_links += len(links)
                document.close()

    # 링크 개수를 출력
    progress_label_download.config(text=f"다운로드 진행: 0/{total_links} 개 링크 다운로드 중")
    
    processed_files = 0
    for root, dirs, files in os.walk(pdf_folder):
        if "files" in dirs:
            dirs.remove("files")
        
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_path = os.path.join(root, file)
                extract_and_download(pdf_path, total_files, processed_files,  0, file_count_label, progress_label_pdf, progress_label_download,progress_label_download_file)
                processed_files += 1

    messagebox.showinfo("완료", "모든 파일 다운로드가 완료되었습니다.")
    start_button.config(text="시작", state=tk.NORMAL)  # 버튼 텍스트 변경 및 활성화

# 폴더 선택 버튼을 눌렀을 때 실행되는 함수
def select_folder():
    global pdf_folder
    pdf_folder = filedialog.askdirectory(initialdir=".", title="PDF 폴더 선택")
    folder_label.config(text=pdf_folder)

# GUI 설정
root = tk.Tk()
root.title("PDF 링크 다운로드 프로그램")

# GUI 스타일 적용
style = ttk.Style()
style.configure("TButton", padding=6, relief="flat", background="#4CAF50", foreground="black")
style.configure("TLabel", font=("Arial", 10), padding=5)

# GUI 구성
frame = tk.Frame(root, bg="#f0f0f0")
frame.pack(padx=20, pady=20)

# 폴더 경로 표시 라벨
folder_label = ttk.Label(frame, text="폴더를 선택하세요", width=50, anchor="w")
folder_label.pack(pady=10)

# 폴더 선택 버튼
select_button = ttk.Button(frame, text="폴더 선택", command=select_folder)
select_button.pack(pady=10)

# 전체 파일 수 표시 라벨
file_count_label = ttk.Label(frame, text="전체 PDF 파일: 0 개", width=50, anchor="w")
file_count_label.pack(pady=10)

# PDF 처리 진행 상태 표시 라벨
progress_label_pdf = ttk.Label(frame, text="진행 중: 0/0 개 PDF 파일 처리 중\n현재 PDF: ", width=50, anchor="w")
progress_label_pdf.pack(pady=10)

# 링크 다운로드 진행 상태 표시 라벨
progress_label_download = ttk.Label(frame, text="다운로드 진행: 0/0 개 링크 다운로드 중", width=50, anchor="w")
progress_label_download.pack(pady=10)

# 링크 다운로드 진행 상태 표시 라벨
progress_label_download_fileName = ttk.Label(frame, text="파일 이름: ", width=50, anchor="w")
progress_label_download_fileName.pack(pady=10)

# 링크 다운로드 진행 상태 표시 라벨
progress_label_download_file = ttk.Label(frame, text="파일 이름: 0/0 개 링크 다운로드 중", width=50, anchor="w")
progress_label_download_file.pack(pady=10)

# 시작 버튼
start_button = ttk.Button(frame, text="시작", command=lambda: process_pdfs(file_count_label, progress_label_pdf, progress_label_download,progress_label_download_file,progress_label_download_fileName))
start_button.pack(pady=10)

# GUI 실행
root.mainloop()
