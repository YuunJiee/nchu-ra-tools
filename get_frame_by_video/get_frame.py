import cv2
import os
from PIL import Image

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
        # 擷取所有影格
        selected_frames = set(range(total_frames))
    else:
        # 均勻選取指定數量的影格
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


def extract_frames_from_folder(root_folder, output_folder=r"Z:\home\114年-114年臺北分署防災及維生通道路網暨農路維護管理與構造物體檢\2_編號農路數化盤點\2_4_frame", frames_to_extract=None):
    for root, _, files in os.walk(root_folder):
        for file in files:
            if file.lower().endswith(".mp4"):
                video_path = os.path.join(root, file)
                extract_frames(video_path, output_folder, frames_to_extract)


if __name__ == "__main__":
    # ✅ 請將這裡改為你要處理的資料夾
    input_folder = r"Z:\home\114年-114年臺北分署防災及維生通道路網暨農路維護管理與構造物體檢\2_編號農路數化盤點\2_2_mp4file\170_農竹新023"
    
    # ✅ 這裡請輸入「要擷取的影格數」，若輸入 0 則擷取所有影格
    frames_to_extract_input = int(input("請輸入要擷取的影格數 (輸入 0 代表全部)："))
    frames_to_extract = frames_to_extract_input if frames_to_extract_input > 0 else None

    extract_frames_from_folder(input_folder, frames_to_extract=frames_to_extract)

