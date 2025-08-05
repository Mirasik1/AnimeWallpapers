import win32con
import win32gui
import win32api
import win32process
import psutil
import threading
import keyboard
from kivy.app import App
from kivy.uix.video import Video
from kivy.core.window import Window
from kivy.clock import Clock
import os
import json
import atexit

SETTINGS_FILE = "settings.json"


class WallpaperApp(App):
    def __init__(self, index=0, **kwargs):
        super().__init__(**kwargs)
        self.index = index

    def build(self):
        with open("video_config.json", "r", encoding="utf-8") as f:
            self.config = json.load(f)

        first_video = self.config["videos"][self.index]
        video_name = first_video["name"]
        self.start_timecode = first_video["timecode"]

        self.video = Video(
            source=f"Videos/{video_name}",
            state='play',
            options={'eos': 'stop'}
        )
        self.video.allow_stretch = True
        self.video.keep_ratio = False
        self.manual_pause = False


        self.video.bind(eos=self.on_video_end)

        Clock.schedule_interval(self.check_fullscreen_window, 1)
        threading.Thread(target=self.listen_keyboard, daemon=True).start()
        Clock.schedule_interval(self.check_video_loaded, 0.1)
        return self.video

    def check_video_loaded(self, dt):
        if self.video._video:
            if self.video._video.duration != -1:
                Clock.unschedule(self.check_video_loaded)
                self.on_video_loaded(self.video)


    def on_video_end(self, instance, value):
        print("video ended")
        self.index += 1
        if self.index >= len(self.config["videos"]):
            print("✅ Все видео проиграны. Зацикливание.")
            self.index = 0
        next_video = self.config["videos"][self.index]
        video_name = next_video["name"]
        self.start_timecode = next_video["timecode"]
        print(next_video, video_name, self.start_timecode)
        self.video.unload()
        self.video.source = f"Videos/{video_name}"
        self.video.state = 'play'
        print(f"▶ Переход к следующему видео: {video_name}")

        Clock.schedule_once(lambda dt: self.seek_video(self.start_timecode), 1.0)

    def on_video_loaded(self, instance):
        Clock.schedule_once(lambda dt: self.seek_video(self.start_timecode), 0.5)



    def seek_video(self, timecode):
        if self.video.duration > 0:
            print(timecode / self.video.duration)
            print(timecode)
            print(self.video.duration)
            self.video.seek(timecode / self.video.duration)
        else:
            Clock.schedule_once(lambda dt: self.seek_video(timecode))

    def seek_forward(self, seconds=10):
        new_position = self.video.position + seconds
        if new_position < self.video.duration:
            self.video.seek(new_position / self.video.duration)
            print(f"⏩ Перемотка вперед на {seconds} сек. Новая позиция: {new_position:.2f} сек.")
        else:
            self.video.seek(1.0)  # конец видео
            print("⏩ Перемотка к концу видео.")

    def seek_backward(self, seconds=10):
        new_position = max(0, self.video.position - seconds)
        self.video.seek(new_position / self.video.duration)
        print(f"⏪ Перемотка назад на {seconds} сек. Новая позиция: {new_position:.2f} сек.")

    def on_start(self):
        Window.clearcolor = (0, 0, 0, 0)
        Window.borderless = True
        Window.top = 0
        Window.left = 0
        Window.size = (win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1))
        Clock.schedule_once(self.set_as_wallpaper, 1)

    def set_as_wallpaper(self, dt):
        progman = win32gui.FindWindow("Progman", None)
        win32gui.SendMessageTimeout(progman, 0x052C, 0xD, 1, win32con.SMTO_NORMAL, 1000)

        def get_workerw():
            workerw = None
            hwnd = win32gui.FindWindowEx(0, None, "WorkerW", None)
            while hwnd:
                next_hwnd = win32gui.FindWindowEx(0, hwnd, "WorkerW", None)
                if next_hwnd:
                    hwnd = next_hwnd
                else:
                    workerw = hwnd
                    break
            return workerw

        workerw = get_workerw()
        if workerw:
            kivy_hwnd = win32gui.GetForegroundWindow()
            win32gui.SetParent(kivy_hwnd, workerw)
            print("Успешно найден WorkerW!")
        else:
            print("Ошибка: не удалось найти окно WorkerW!")

    def check_fullscreen_window(self, dt):
        if is_fullscreen_window():
            if self.video.state == 'play':
                self.video.state = 'pause'

                print("🔴 Полноэкранное окно найдено. Видео на паузе.")
        else:
            if self.video.state == 'pause' and not self.manual_pause:
                self.save_timecode(self.video.source, self.video.position)
                self.video.state = 'play'
                print("🟢 Видео продолжает воспроизведение.")

    def on_stop(self):
        print("Приложение завершилось корректно.")
        if self.video:
            self.save_timecode(self.video.source, self.video.position)
            print("📌 Сохранение таймкода перед закрытием:", self.video.position)

    def listen_keyboard(self):
        keyboard.add_hotkey("ctrl+shift+s", lambda: self.toggle_play_pause())
        keyboard.add_hotkey("ctrl+shift+right", lambda: self.seek_forward())
        keyboard.add_hotkey("ctrl+shift+left", lambda: self.seek_backward())
        keyboard.wait()

    def save_timecode(self, video_name, timecode):
        for video in self.config["videos"]:
            if video["name"] == video_name.split("/")[-1]:
                prev_timecode = video.get("timecode", 0)  # Получаем предыдущий таймкод, по умолчанию 0
                if timecode > prev_timecode:  # Записываем только если текущий таймкод больше предыдущего
                    video["timecode"] = timecode
                    with open("video_config.json", "w", encoding="utf-8") as f:
                        json.dump(self.config, f, indent=4)
                    print("✅ Таймкод сохранён:", timecode)
                else:
                    print("ℹ Новый таймкод меньше или равен предыдущему. Сохранение не требуется.")
                break

    def toggle_play_pause(self):
        if is_fullscreen_window():
            print("⚠ Невозможно включить видео: найдено полноэкранное окно.")
            return

        if self.video.state == 'play':
            self.save_timecode(self.video.source, self.video.position)
            self.video.state = 'pause'
            self.manual_pause = True
            print("⏸ Видео на паузе.")

        else:
            self.video.state = 'play'
            self.manual_pause = False
            print(f"▶ Видео воспроизводится с {self.start_timecode} сек.")


def create_settings():
    if not os.path.exists(SETTINGS_FILE):
        settings = {
            "config_file": "video_config.json",
            "videos_folder": "Videos"
        }
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=4)
        print("🆕 Файл settings.json создан!")
    else:
        print("✅ settings.json уже существует.")


def create_video_config():
    create_settings()

    with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
        settings = json.load(f)

    config_file = settings["config_file"]
    videos_folder = settings["videos_folder"]

    # Список всех видеофайлов в папке
    current_files = [
        file for file in os.listdir(videos_folder)
        if file.lower().endswith((".mp4", ".avi", ".mov", ".mkv"))
    ] if os.path.exists(videos_folder) else []

    # Загрузка существующего файла конфигурации
    if os.path.exists(config_file):
        with open(config_file, "r", encoding="utf-8") as f:
            config_data = json.load(f)

        existing_videos = config_data.get("videos", [])
        existing_names = [video["name"] for video in existing_videos]
        current_index = config_data.get("current_index", 0)

        # Удаление несуществующих файлов
        updated_videos = [
            video for video in existing_videos if video["name"] in current_files
        ]

        # Добавление новых файлов
        for file in current_files:
            if file not in existing_names:
                updated_videos.append({"name": file, "timecode": 0})

        config_data = {
            "videos": updated_videos,
            "current_index": min(current_index, len(updated_videos) - 1)
        }

        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=4)

        print("🔁 Конфигурация обновлена.")
    else:
        # Если файла конфигурации ещё нет
        videos = [{"name": file, "timecode": 0} for file in current_files]
        config_data = {"videos": videos, "current_index": 0}

        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=4)

        print(f"🆕 Файл {config_file} создан!")


def is_fullscreen_window():
    fg_window = win32gui.GetForegroundWindow()
    if not fg_window:
        return False

    rect = win32gui.GetWindowRect(fg_window)
    screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

    placement = win32gui.GetWindowPlacement(fg_window)
    is_maximized = placement[1] == win32con.SW_SHOWMAXIMIZED

    is_fullscreen = (rect[0] <= 0 and rect[1] <= 0 and
                     rect[2] >= screen_width and rect[3] >= screen_height)

    _, pid = win32process.GetWindowThreadProcessId(fg_window)
    process_name = psutil.Process(pid).name().lower()

    if "explorer.exe" in process_name:
        return False

    return is_maximized or is_fullscreen


def save_on_exit():
    if hasattr(app, 'video') and app.video:
        app.save_timecode(app.video.source, app.video.position)
        print("📌 Сохранение таймкода при выходе:", app.video.position)

        try:
            with open("video_config.json", "r", encoding="utf-8") as f:
                config = json.load(f)
            config["current_index"] = app.index
            with open("video_config.json", "w", encoding="utf-8") as f:
                json.dump(config, f, indent=4)
            print("💾 Сохранён current_index:", app.index)
        except Exception as e:
            print("⚠ Ошибка при сохранении current_index:", e)


if __name__ == "__main__":
    create_video_config()
    with open("video_config.json", "r", encoding="utf-8") as f:
        config = json.load(f)
        current_index = config.get("current_index", 0)

    app = WallpaperApp(index=current_index)
    atexit.register(save_on_exit)
    app.run()
