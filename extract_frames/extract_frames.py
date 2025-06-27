import cv2
import os
from PIL import Image
from tkinter import Tk, filedialog

def extract_frames(video_path, output_folder="output", frames_to_extract=None):
    video_name = os.path.splitext(os.path.basename(video_path))[0]
    save_dir = os.path.join(output_folder, video_name)
    os.makedirs(save_dir, exist_ok=True)

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        print(f"❌ 無法開啟影片：{video_path}")
        return

    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    if total_frames == 0:
        print(f"⚠️ 無影格：{video_path}")
        return

    print(f"🎞️ {video_path} - 總影格數: {total_frames}")

    if frames_to_extract is None or frames_to_extract >= total_frames:
        selected_frames = set(range(total_frames))
    else:
        selected_frames = set(int(i * total_frames / frames_to_extract) for i in range(frames_to_extract))

    frame_idx = 0
    saved_count = 0

    while cap.isOpened():
        ret, frame = cap.read()
        if not ret:
            break

        if frame_idx in selected_frames:
            frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            image = Image.fromarray(frame_rgb)
            filename = os.path.join(save_dir, f"{video_name}_{frame_idx:05d}.png")
            image.save(filename)
            saved_count += 1
            print(f"✅ 儲存影格 {frame_idx} 至 {filename}")

        frame_idx += 1

    cap.release()
    print(f"🎉 完成：{video_path} → 共儲存 {saved_count} 張影格\n")


if __name__ == "__main__":
    # 關閉 tkinter 主視窗
    root = Tk()
    root.withdraw()

    # ✅ 選擇影片檔
    video_path = filedialog.askopenfilename(
        title="請選擇影片檔（.mp4）",
        filetypes=[("MP4 Video", "*.mp4")]
    )
    if not video_path:
        print("⚠️ 未選擇影片，程式結束。")
        exit()

    # ✅ 選擇輸出資料夾
    output_folder = filedialog.askdirectory(
        title="請選擇輸出資料夾"
    )
    if not output_folder:
        print("⚠️ 未選擇輸出資料夾，程式結束。")
        exit()

    # ✅ 輸入擷取影格數
    try:
        frames_to_extract_input = int(input("請輸入要擷取的影格數 (輸入 0 代表全部)："))
        frames_to_extract = frames_to_extract_input if frames_to_extract_input > 0 else None
    except:
        print("⚠️ 輸入格式錯誤，預設擷取全部影格。")
        frames_to_extract = None

    # ✅ 執行擷取
    extract_frames(video_path, output_folder, frames_to_extract)

