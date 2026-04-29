import os
import re
import glob
import time
import win32com.client

# Constants for PowerPoint
try:
    PP_EFFECT_PUSH_UP = win32com.client.constants.ppEffectPushUp
except Exception:
    # Fallback: Push Up transition
    PP_EFFECT_PUSH_UP = 3855
try:
    MSO_ANIM_TRIGGER_WITH_PREVIOUS = win32com.client.constants.MSO_ANIM_TRIGGER_WITH_PREVIOUS
except Exception:
    MSO_ANIM_TRIGGER_WITH_PREVIOUS = getattr(win32com.client.constants, 'msoAnimTriggerWithPrevious', 2)

try:
    PP_MEDIA_TASK_STATUS_IN_PROGRESS = win32com.client.constants.ppMediaTaskStatusInProgress
except Exception:
    PP_MEDIA_TASK_STATUS_IN_PROGRESS = 1
try:
    PP_MEDIA_TASK_STATUS_QUEUED = win32com.client.constants.ppMediaTaskStatusQueued
except Exception:
    PP_MEDIA_TASK_STATUS_QUEUED = 2
try:
    PP_MEDIA_TASK_STATUS_DONE = win32com.client.constants.ppMediaTaskStatusDone
except Exception:
    PP_MEDIA_TASK_STATUS_DONE = 3
try:
    PP_MEDIA_TASK_STATUS_FAILED = win32com.client.constants.ppMediaTaskStatusFailed
except Exception:
    PP_MEDIA_TASK_STATUS_FAILED = 4

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]


def wait_for_video_export(presentation):
    while True:
        status = presentation.CreateVideoStatus
        if status in (
            PP_MEDIA_TASK_STATUS_DONE,
            PP_MEDIA_TASK_STATUS_FAILED,
        ):
            return status
        if status not in (PP_MEDIA_TASK_STATUS_IN_PROGRESS, PP_MEDIA_TASK_STATUS_QUEUED):
            return status
        time.sleep(2)

def main():
    pp_app = None
    try:
        pp_app = win32com.client.Dispatch("PowerPoint.Application")
        pp_app.Visible = False
    except Exception as e:
        print(f"Error: {e}")
        return

    presentation = None
    try:
        current_dir = os.getcwd()
        all_pptx = glob.glob(os.path.join(current_dir, "*.pptx"))
        valid_files = [f for f in all_pptx if not f.endswith('_int.pptx')]

        if not valid_files:
            print("No input .pptx files found.")
            return

        input_path = os.path.abspath(valid_files[0])
        output_path = input_path.replace(".pptx", "_int.pptx")
        video_output_path = input_path.replace(".pptx", "_video.mp4")
        presentation = pp_app.Presentations.Open(input_path)

        audio_files = [f for f in os.listdir('.') if f.endswith('.wav')]
        audio_files.sort(key=natural_sort_key)

        for i, slide in enumerate(presentation.Slides, start=1):
            slide.SlideShowTransition.EntryEffect = PP_EFFECT_PUSH_UP

            if (i - 1) >= len(audio_files):
                continue

            for j in range(slide.Shapes.Count, 0, -1):
                if slide.Shapes(j).Type == 16:
                    slide.Shapes(j).Delete()

            audio_full_path = os.path.abspath(audio_files[i - 1])
            icon_size = 18
            left = presentation.PageSetup.SlideWidth - icon_size - 100
            top = presentation.PageSetup.SlideHeight - icon_size - 10

            try:
                shape = slide.Shapes.AddMediaObject2(audio_full_path, False, True, left, top, icon_size, icon_size)
                play_effect = slide.TimeLine.MainSequence.AddEffect(shape, 83)
                play_effect.Timing.TriggerType = MSO_ANIM_TRIGGER_WITH_PREVIOUS
                shape.AnimationSettings.PlaySettings.HideWhileNotPlaying = True

                print(f"Slide {i}: transition set and audio {audio_files[i - 1]} added with previous")
            except Exception as e:
                print(f"Slide {i} error: {e}")

        presentation.SaveAs(output_path)

        try:
            presentation.CreateVideo(video_output_path, True, 5, 720, 30, 85)
            status = wait_for_video_export(presentation)
            if status == PP_MEDIA_TASK_STATUS_DONE:
                print(f"Video exported successfully: {video_output_path}")
            else:
                print(f"Video export finished with status {status}")
        except Exception as e:
            print(f"Video export skipped: {e}")

        print(f"Intermediate presentation saved as: {output_path}")
        print("Process finished successfully.")
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if pp_app is not None:
            try:
                pp_app.Quit()
            except Exception:
                pass

if __name__ == "__main__":
    main()
