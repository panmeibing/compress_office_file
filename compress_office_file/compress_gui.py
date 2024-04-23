import os.path
import shutil
import time
from multiprocessing.pool import ThreadPool
from threading import Thread
from tkinter.filedialog import askdirectory, askopenfilenames
from tkinter.messagebox import askyesno, showwarning

import ttkbootstrap as ttkb
from ttkbootstrap.tooltip import ToolTip

from utils.image_utils import compress_local_image
from utils.logger_utils import get_logger
from utils.zip_utils import unzip_file, zip_file

FILE_TYPE_MAP = {
    "Word": ".docx",
    "Excel": ".xlsx",
    "PowerPoint": ".pptx"
}

COMPRESS_IMG_TYPE = {"png", "jpg", "jpeg", "bmp"}
MAX_THREAD_COUNT = 20
COMPRESS_BTN_TEXT = " 开始压缩 "
COMPRESS_BTN_DISABLE_TEXT = "处理中..."

ROOT_PATH = os.path.dirname(__file__)
print("ROOT_PATH: ", ROOT_PATH)

logger = get_logger(os.path.join(ROOT_PATH, "logs"), "compress.log")

logger.info("init_log")


class CompressOfficeFile:

    def __init__(self, master=None):
        self.root = master if master else ttkb.Window(title="Office文档压缩", resizable=(False, False))
        # self.root.iconbitmap('')
        self.tmp_dir = os.path.join(ROOT_PATH, "tmp")
        self.container_frame = ttkb.Frame(self.root, padding=(100, 10, 100, 30))
        self.main_frame = ttkb.Frame(self.container_frame)
        self.params_frame = ttkb.Frame(self.main_frame)
        self.logs_frame = ttkb.Frame(self.main_frame)
        self.target_save_dir_var = ttkb.StringVar()
        self.todo_files = ttkb.StringVar()
        self.text_entry = None
        self.compress_button = None
        self.tip_label = None
        self.tip_label_var = ttkb.StringVar(value="")
        self.create_view()
        self.center_window()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()

    def create_view(self):
        logger.info("create_view()")
        entry_width = 50
        entry_pady = (7, 18)
        # 标题
        ttkb.Label(self.root, text="Office文档压缩工具", font=(None, 20, "bold")).pack(pady=50)
        # 选择要处理的文件
        to_do_files_lb = ttkb.Label(self.params_frame, text='待处理文件')
        to_do_files_lb.pack(anchor="w")
        ToolTip(to_do_files_lb, text="请选择要处理的docx、xlsx或pptx文件，可以多选")
        todo_files_frame = ttkb.Frame(self.params_frame)
        self.text_entry = ttkb.Text(todo_files_frame, width=entry_width, height=1)
        self.text_entry.pack(side='left', fill='x', expand=1, pady=entry_pady)
        ttkb.Button(todo_files_frame, text='浏览', command=self.on_select_files).pack(
            side="right", padx=(10, 0), fill="x", pady=entry_pady)
        todo_files_frame.pack(side="top", fill="x", expand=1)
        # 保存路径，点击选取一个文件夹
        save_path_lb = ttkb.Label(self.params_frame, text="保存路径")
        save_path_lb.pack(anchor="w")
        ToolTip(save_path_lb, text="要保存的文件夹，请确保路径合法并存在")
        save_dir_frame = ttkb.Frame(self.params_frame)
        ttkb.Entry(save_dir_frame, textvariable=self.target_save_dir_var).pack(
            side='left', fill="x", expand=1, pady=entry_pady)
        ttkb.Button(save_dir_frame, text='浏览', command=lambda: self.target_save_dir_var.set(askdirectory())).pack(
            side="right", padx=(10, 0), fill="x", pady=entry_pady)
        save_dir_frame.pack(side="top", fill="x", expand=1)
        # 提示信息
        self.tip_label = ttkb.Label(self.params_frame, textvariable=self.tip_label_var, bootstyle="secondary")
        self.tip_label.pack(side="top", fill="x", expand=1)
        # 显示各种Frame
        self.params_frame.grid(row=0, column=0, padx=20)
        self.main_frame.pack(pady=10)
        # 压缩按钮
        self.compress_button = ttkb.Button(self.container_frame, text=COMPRESS_BTN_TEXT, command=self.on_compress_job)
        self.compress_button.pack(pady=20)
        self.container_frame.pack()
        ttkb.Label(self.root, text="made by 冰冷的希望", font=(None, 10), bootstyle="secondary").pack(pady=10)

    def center_window(self):
        logger.info("center_window()")
        try:
            self.root.update()
            width, height = self.root.winfo_width(), self.root.winfo_height()
            center_width, = (self.root.winfo_screenwidth() - width) // 2,
            center_height = (self.root.winfo_screenheight() - height) // 2
            geometry_str = "+{}+{}".format(center_width, center_height)
            self.root.geometry(geometry_str)
        except Exception as e:
            print(f"center_window() error: {e}")

    def on_closing(self):
        self.root.destroy()
        # result = askyesno("提示", "确定要退出吗？")
        # if result:
        #     self.root.destroy()

    def on_select_files(self):
        selected_res = askopenfilenames(filetypes=[(k, v) for k, v in FILE_TYPE_MAP.items()])
        self.text_entry.delete("1.0", ttkb.END)
        self.text_entry.edit_reset()
        print("selected_res: ", selected_res)
        height_line = len(selected_res) if selected_res else 1
        self.text_entry.config(height=height_line)
        self.text_entry.insert("1.0", "\n".join(selected_res))

    def on_compress_job(self):
        t = Thread(target=self.start_compress_job)
        t.start()

    def start_compress_job(self):
        print("start_compress_job()")
        is_ok, path_data = self.check_params()
        if not is_ok:
            logger.warning(f"start_compress_job() check params is not ok: {path_data}")
            showwarning(title="提示", message=path_data)
            return
        total_file_count = len(path_data)
        print(f"start_compress_job() total file count: {total_file_count}")
        self.compress_button.config(text=COMPRESS_BTN_DISABLE_TEXT)
        self.compress_button.config(state="disabled")
        success_count = 0
        for file_index, file_path in enumerate(path_data):
            do_compress_tip = f"压缩中...  {file_index + 1}/{total_file_count} file_path: {file_path}"
            self.tip_label_var.set(do_compress_tip)
            print(f"start_compress_job(): {do_compress_tip}")
            start_time = time.time()
            try:
                self.compress(file_path)
                success_count += 1
            except Exception as compress_error:
                logger.error(f"start_compress_job() compress error: {compress_error}")
            print(f"start_compress_job() file_path {file_index} spent time: {time.time() - start_time}")
        self.tip_label_var.set(f"已压缩，共 {total_file_count} 个，成功 {success_count} 个")
        self.compress_button.config(text=COMPRESS_BTN_TEXT)
        self.compress_button.config(state='normal')

    def compress(self, file_path):
        logger.info(f"compress() start, file_path: {file_path}")
        file_name = os.path.split(file_path)[1]
        tmp_unzip_dir = os.path.join(self.tmp_dir, f"{file_name}_{int(time.time())}")
        logger.info(f"compress() tmp_unzip_dir: {tmp_unzip_dir}")
        if not os.path.exists(tmp_unzip_dir):
            os.makedirs(tmp_unzip_dir, exist_ok=True)
        try:
            unzip_file(file_path, tmp_unzip_dir)
        except Exception as e:
            logger.error(f"compress() unzip file error: {e}")
            return
        file_type = os.path.splitext(file_path)[1]
        if file_type == ".pptx":
            self.compress_img_multi_threads(os.path.join(tmp_unzip_dir, "ppt", "media"))
        elif file_type == ".docx":
            self.compress_img_multi_threads(os.path.join(tmp_unzip_dir, "word", "media"))
        elif file_type == ".xlsx":
            self.compress_img_multi_threads(os.path.join(tmp_unzip_dir, "xl", "media"))
        else:
            logger.warning(f"compress() unsupported file_type: {file_type}")
        target_save_dir = self.target_save_dir_var.get()
        target_zipped_file = os.path.join(target_save_dir, f"{file_name.split('.')[0]}_compressed{file_type}")
        logger.info(f"compress() target_zipped_file: {target_zipped_file}")
        try:
            zip_file(tmp_unzip_dir, target_zipped_file)
        except Exception as e:
            logger.error(f"compress() zip file error: {e}")
        try:
            # del_all_files(tmp_unzip_dir)
            shutil.rmtree(tmp_unzip_dir)
        except Exception as e:
            logger.error(f"compress() remove tmp dirs failed: {e}")
        logger.info(f"compress() finished, file_path: {file_path}")

    def check_params(self):
        files_str = self.text_entry.get("1.0", ttkb.END).strip()
        logger.info(f"check_params() files_str: {files_str}")
        if not files_str:
            return False, "请选择要压缩的Office文件"
        file_path_list = files_str.split("\n")
        for file_path in file_path_list:
            if not os.path.exists(file_path):
                return False, f"文件不存在：{file_path}"
            if not os.path.isfile(file_path):
                return False, f"路径不是文件：{file_path}"
            file_ext_name = os.path.splitext(file_path)[1]
            if file_ext_name not in FILE_TYPE_MAP.values():
                return False, f"不支持的文件类型: {file_path}"
        target_save_dir = self.target_save_dir_var.get()
        if not target_save_dir or not os.path.exists(target_save_dir) or not os.path.isdir(target_save_dir):
            return False, "请选择一个合法的保存文件夹"
        return True, file_path_list

    def compress_img_multi_threads(self, img_path):
        if not os.path.exists(img_path):
            return
        pool = ThreadPool(MAX_THREAD_COUNT)
        for file_name in os.listdir(img_path):
            file_type = os.path.splitext(file_name)[1]  # .png
            ext_type = file_type.strip().replace(".", "").lower()
            if ext_type not in COMPRESS_IMG_TYPE:
                continue
            img_full_path = os.path.join(img_path, file_name)
            pool.apply_async(compress_local_image, args=(img_full_path,))
        pool.close()
        pool.join()


def run_gui():
    CompressOfficeFile()


if __name__ == '__main__':
    run_gui()
