import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import fitz
from datetime import datetime

# PDFファイルのフルパス一覧
pdf_files = []

# 拡大率
zoom_scale = 1.0

# 現在表示中のPDF
current_pdf_path = None

# Canvas表示用の画像を保持
current_photo = None


# -------------------------
# ログ表示
# -------------------------
def write_log(message):
    now = datetime.now().strftime("%H:%M:%S")
    log_text.insert(tk.END, f"[{now}] {message}\n")
    log_text.see(tk.END)


# -------------------------
# PDFプレビュー表示
# -------------------------
def show_pdf_preview(pdf_path):
    global current_pdf_path, current_photo
    current_pdf_path = pdf_path

    try:
        doc = fitz.open(pdf_path)
        page = doc[0]

        # 拡大率を反映
        mat = fitz.Matrix(zoom_scale, zoom_scale)
        pix = page.get_pixmap(matrix=mat)

        # Pillow画像に変換
        if pix.alpha:
            image = Image.frombytes("RGBA", (pix.width, pix.height), pix.samples)
        else:
            image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

        current_photo = ImageTk.PhotoImage(image)

        preview_canvas.delete("all")
        preview_canvas.create_image(0, 0, anchor="nw", image=current_photo)
        preview_canvas.config(scrollregion=preview_canvas.bbox("all"))

        doc.close()

    except Exception as e:
        preview_canvas.delete("all")
        preview_canvas.create_text(
            250, 100,
            text=f"プレビュー表示エラー\n{e}",
            anchor="center"
        )
        write_log(f"プレビュー表示エラー: {e}")


# -------------------------
# 指定したindexのPDFを選択して表示
# -------------------------
def select_pdf_by_index(index):
    if index < 0 or index >= len(pdf_files):
        return

    file_listbox.selection_clear(0, tk.END)
    file_listbox.selection_set(index)
    file_listbox.activate(index)
    file_listbox.see(index)

    pdf_path = pdf_files[index]

    filename = os.path.basename(pdf_path)
    name_only = os.path.splitext(filename)[0]

    filename_entry.delete(0, tk.END)
    filename_entry.insert(0, name_only)

    show_pdf_preview(pdf_path)


# -------------------------
# リストでファイル選択したとき
# -------------------------
def on_file_select(event):
    selected_index = file_listbox.curselection()
    if not selected_index:
        return

    index = selected_index[0]
    pdf_path = pdf_files[index]

    filename = os.path.basename(pdf_path)
    name_only = os.path.splitext(filename)[0]

    filename_entry.delete(0, tk.END)
    filename_entry.insert(0, name_only)

    show_pdf_preview(pdf_path)


# -------------------------
# フォルダ選択
# -------------------------
def select_folder():
    global zoom_scale, current_pdf_path

    folder_path = filedialog.askdirectory(title="フォルダを選択してください")
    if not folder_path:
        return

    file_listbox.delete(0, tk.END)
    pdf_files.clear()
    zoom_scale = 1.0
    current_pdf_path = None
    preview_canvas.delete("all")

    for file_name in os.listdir(folder_path):
        if file_name.lower().endswith(".pdf"):
            full_path = os.path.join(folder_path, file_name)
            pdf_files.append(full_path)
            file_listbox.insert(tk.END, file_name)

    if not pdf_files:
        preview_canvas.create_text(
            250, 100,
            text="PDFファイルが見つかりません",
            anchor="center"
        )
        write_log("PDFファイルが見つかりませんでした。")
    else:
        preview_canvas.create_text(
            250, 100,
            text="PDFを選択するとここに表示されます",
            anchor="center"
        )
        write_log(f"フォルダを読み込みました。PDF {len(pdf_files)} 件")


# -------------------------
# 拡大
# -------------------------
def zoom_in():
    global zoom_scale
    zoom_scale *= 1.2

    if current_pdf_path:
        show_pdf_preview(current_pdf_path)
        write_log(f"拡大しました（倍率: {zoom_scale:.2f}）")


# -------------------------
# 縮小
# -------------------------
def zoom_out():
    global zoom_scale
    zoom_scale /= 1.2

    if current_pdf_path:
        show_pdf_preview(current_pdf_path)
        write_log(f"縮小しました（倍率: {zoom_scale:.2f}）")


# -------------------------
# リネーム処理
# -------------------------
def rename_file(event=None):
    selected_index = file_listbox.curselection()

    if not selected_index:
        messagebox.showwarning("確認", "先にPDFファイルを選択してください。")
        write_log("リネーム失敗: ファイルが選択されていません。")
        return

    new_name = filename_entry.get().strip()

    if not new_name:
        messagebox.showwarning("確認", "新しいファイル名を入力してください。")
        write_log("リネーム失敗: 新しいファイル名が空です。")
        return

    ng_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    for ch in ng_chars:
        if ch in new_name:
            messagebox.showerror(
                "エラー",
                f"ファイル名に使えない文字が含まれています。\n{ch}"
            )
            write_log(f"リネーム失敗: 使用できない文字 '{ch}' が含まれています。")
            return

    try:
        index = selected_index[0]
        old_path = pdf_files[index]
        old_filename = os.path.basename(old_path)

        folder_path = os.path.dirname(old_path)
        new_filename = new_name + ".pdf"
        new_path = os.path.join(folder_path, new_filename)

        if os.path.exists(new_path) and old_path != new_path:
            messagebox.showerror("エラー", "同じ名前のファイルがすでに存在します。")
            write_log(f"リネーム失敗: '{new_filename}' はすでに存在します。")
            return

        os.rename(old_path, new_path)

        pdf_files[index] = new_path
        file_listbox.delete(index)
        file_listbox.insert(index, new_filename)

        write_log(f"リネーム成功: {old_filename} → {new_filename}")

        # 次のファイルへ自動移動
        if len(pdf_files) == 0:
            return

        next_index = index + 1

        # 最後のファイルだった場合は、そのまま現在の行を選択
        if next_index >= len(pdf_files):
            next_index = index

        select_pdf_by_index(next_index)

    except Exception as e:
        messagebox.showerror("エラー", f"リネームに失敗しました。\n{e}")
        write_log(f"リネーム失敗: {e}")


# =========================
# メインウィンドウ
# =========================
root = tk.Tk()
root.title("PDFリネームツール")
root.geometry("1100x800")

main_frame = ttk.Frame(root)
main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

# 上部ボタンエリア
top_button_frame = ttk.Frame(main_frame)
top_button_frame.pack(fill=tk.X, pady=(0, 10))

folder_button = ttk.Button(
    top_button_frame,
    text="フォルダを選択",
    command=select_folder
)
folder_button.pack(side=tk.LEFT)

# 中央エリア
top_frame = ttk.Frame(main_frame)
top_frame.pack(fill=tk.BOTH, expand=True)

# 下部入力エリア
bottom_frame = ttk.Frame(main_frame)
bottom_frame.pack(fill=tk.X, pady=10)

# ログエリア
log_frame = ttk.Frame(main_frame)
log_frame.pack(fill=tk.BOTH, expand=False)

# -------------------------
# 左：ファイル一覧
# -------------------------
left_frame = ttk.Frame(top_frame)
left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 5))

label_left = ttk.Label(left_frame, text="PDFファイル一覧")
label_left.pack(anchor="w")

file_listbox = tk.Listbox(left_frame, width=40)
file_listbox.pack(fill=tk.BOTH, expand=True)
file_listbox.bind("<<ListboxSelect>>", on_file_select)

# -------------------------
# 右：プレビュー
# -------------------------
right_frame = ttk.Frame(top_frame)
right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

label_right = ttk.Label(right_frame, text="PDFプレビュー")
label_right.pack(anchor="w")

zoom_frame = ttk.Frame(right_frame)
zoom_frame.pack(fill=tk.X, pady=5)

zoom_in_button = ttk.Button(zoom_frame, text="拡大", command=zoom_in)
zoom_in_button.pack(side=tk.LEFT, padx=5)

zoom_out_button = ttk.Button(zoom_frame, text="縮小", command=zoom_out)
zoom_out_button.pack(side=tk.LEFT)

canvas_frame = ttk.Frame(right_frame)
canvas_frame.pack(fill=tk.BOTH, expand=True)

x_scroll = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL)
y_scroll = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL)

preview_canvas = tk.Canvas(
    canvas_frame,
    bg="lightgray",
    xscrollcommand=x_scroll.set,
    yscrollcommand=y_scroll.set
)

x_scroll.config(command=preview_canvas.xview)
y_scroll.config(command=preview_canvas.yview)

x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

preview_canvas.create_text(
    250, 100,
    text="フォルダを選択してください",
    anchor="center"
)

# -------------------------
# 下：ファイル名入力＋ボタン
# -------------------------
label_entry = ttk.Label(bottom_frame, text="新しいファイル名:")
label_entry.pack(side=tk.LEFT)

filename_entry = ttk.Entry(bottom_frame)
filename_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
filename_entry.bind("<Return>", rename_file)

rename_button = ttk.Button(bottom_frame, text="リネーム", command=rename_file)
rename_button.pack(side=tk.RIGHT)

# -------------------------
# ログ表示欄
# -------------------------
label_log = ttk.Label(log_frame, text="処理ログ")
label_log.pack(anchor="w")

log_text = tk.Text(log_frame, height=8)
log_text.pack(fill=tk.BOTH, expand=True)

write_log("アプリを起動しました。")

root.mainloop()