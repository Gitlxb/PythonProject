# -- coding: utf-8 --
# @Time : 2025-09-04 15:32
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : sc_spzwz_wl.py
# @Software: PyCharm

import moviepy.editor as mp
import whisper


# 步骤 1: 从视频提取音频
def extract_audio_from_video(video_path, output_audio_path):
    # 加载视频文件
    clip = mp.VideoFileClip(video_path)
    # 提取音频并保存为wav文件
    clip.audio.write_audiofile(output_audio_path)
    # 关闭视频流，释放资源
    clip.close()


# 步骤 2: 使用Whisper识别音频
def transcribe_audio(audio_path):
    # 加载模型，默认是 'small' 模型，精度和速度平衡
    # 可选模型: 'tiny', 'base', 'small', 'medium', 'large'
    # 模型越大越精确，但速度越慢
    model = whisper.load_model("small")

    # 转录音频
    result = model.transcribe(audio_path, language='zh')  # 如果视频是中文，指定语言可提高准确性

    # 返回识别结果
    return result["text"]


# 主程序
if __name__ == "__main__":
    video_file = "你的视频文件路径.mp4"
    audio_file = "提取的音频.wav"

    print("正在从视频中提取音频...")
    extract_audio_from_video(video_file, audio_file)

    print("正在将音频转换为文字...")
    text = transcribe_audio(audio_file)

    print("识别结果：")
    print(text)

    # 你也可以将结果保存到文件
    with open("转录文本.txt", "w", encoding="utf-8") as f:
        f.write(text)