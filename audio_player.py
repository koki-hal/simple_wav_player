import multiprocessing
import pyaudio
import wave
import time


class Play:
    PLAY = 1
    PAUSE = 2
    STOP = 0


class Playing:
    PLAYING = 1
    FINISH = 0


def _play_audio(device, wav_file, play, playing, event):
    """
    Play an audio file executed by a process.

    ATTENTION:
        PyAudio (based on PortAudio) is not thread-safe.
        Also, Python is running under GIL.
        If this function is running as a thread, it won't work correctly.
    """

    # print('Start Process...') # _FOR_DEBUG_
    wf = wave.open(wav_file, 'rb')
    sw = wf.getsampwidth()
    ch = wf.getnchannels()
    fr = wf.getframerate()

    p = pyaudio.PyAudio()
    fmt = p.get_format_from_width(sw)

    # Check the channel count
    wav_channels = ch
    device_channels = int(device['maxOutputChannels'])
    output_channels = min(wav_channels, device_channels)

    # Adjust the sampling rate regarding the channel count
    fr = int(fr * wav_channels / output_channels)

    stream = p.open(
        format=fmt,
        channels=output_channels,
        rate=fr,
        output=True,
        output_device_index=device['index'],
    )
    
    chunk = 2 ** 10
    data = wf.readframes(chunk)
    stream_available = True
    # print('Playing...') # _FOR_DEBUG_
    while data:
        try:
            event.wait()

            if play.value == Play.PLAY:
                pass
            elif play.value == Play.PAUSE:
                # continue
                pass
            elif play.value == Play.STOP:
                break
            
            stream.write(data)
            data = wf.readframes(chunk)
        except OSError as e:
            # stream can't be used anymore
            # possibly, the device is disconnected before finish playing
            stream_available = False
            break

    # print('Finished Playing...') # _FOR_DEBUG_

    playing.value = Playing.FINISH

    if stream_available:
        stream.stop_stream()
        stream.close()
    p.terminate()

    # print('Exit Playing...') # _FOR_DEBUG_


class AudioPlayer:
    def __init__(self):
        self.play = multiprocessing.Value('i', Play.STOP)
        self.playing = multiprocessing.Value('i', Playing.FINISH)
        self.play_process = None
        self.event = multiprocessing.Event()

    def play_audio(self, device_name, wav_file):
        """
        Play an audio file.
        
        Args:
            device_name (str): The friendly name of the audio device.
            wav_file (str): The path of the WAV file.
        """

        if self.play_process is None:
            # Playing process is not started.
            device = self._get_device(device_name)
            if device is None:
                # print('Device not found.')
                return

            self.play.value = Play.PLAY
            self.playing.value = Playing.PLAYING
            self.event.set()
            self.play_process = multiprocessing.Process(target=_play_audio, args=(device, wav_file, self.play, self.playing, self.event))
            self.play_process.start()
        else:
            # PAUSE
            self.play.value = Play.PLAY
            self.event.set()

    def pause_audio(self):
        self.play.value = Play.PAUSE
        self.event.clear()

    def stop_audio(self):
        self.play.value = Play.STOP
        self.event.set()
        while self.is_playing:
            time.sleep(0.1)
        self.play_process = None

    def audio_finished(self):
        # If the audio is finished naturally, the process is finished but the instance variable is not cleared.
        # In this case, this method is needed to be called just to clear the variable.
        self.play_process = None
        self.event.set()

    @property
    def is_playing(self):
        return self.playing.value == Playing.PLAYING

    def _get_device(self, device_friendly_name):
        """
        Return the PyAudio object regarding the device friendly name.
        """

        p = pyaudio.PyAudio()
        host_api_count = p.get_host_api_count()
        for i in range(host_api_count):
            host_api = p.get_host_api_info_by_index(i)
            if host_api['name'] == 'MME':
                device_count = int(host_api['deviceCount'])
                for j in range(device_count):
                    device = p.get_device_info_by_host_api_device_index(i, j)
                    if int(device['maxOutputChannels']) > 0:
                        name = str(device['name'])
                        if name == device_friendly_name[:len(name)]:
                            p.terminate()
                            return device
        p.terminate()
        return None


