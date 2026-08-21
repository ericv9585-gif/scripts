import sys
import time
import queue
import threading

import pythoncom
import win32com.client

from selenium import webdriver
from selenium.webdriver.edge.options import Options

from PySide6.QtCore import Qt, QObject, Signal
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QListWidget,
    QSlider,
    QFrame,
)


# ============================================================
# SETTINGS
# ============================================================

POLL_INTERVAL = 0.15

DEFAULT_RATE = 0
MIN_RATE = -10
MAX_RATE = 10


# ============================================================
# TWITCH JAVASCRIPT
# ============================================================

OBSERVER_JS = r"""
(function () {

    if (window.__twitchVoiceInstalled) {
        return;
    }

    window.__twitchVoiceInstalled = true;
    window.__twitchVoiceQueue = [];

    function parseMessage(element) {

        if (!element) {
            return null;
        }

        var usernameElement =
            element.querySelector(
                '[data-a-target="chat-message-username"]'
            );

        var textElement =
            element.querySelector(
                '[data-a-target="chat-message-text"]'
            );

        if (!usernameElement || !textElement) {
            return null;
        }

        var username =
            usernameElement.textContent.trim();

        var text =
            textElement.textContent.trim();

        if (!username || !text) {
            return null;
        }

        return {
            username: username,
            text: text
        };
    }


    function inspectNode(node) {

        if (!node || node.nodeType !== 1) {
            return;
        }

        if (
            node.matches(
                '[data-a-target="chat-line-message"]'
            )
        ) {
            var message = parseMessage(node);

            if (message) {
                window.__twitchVoiceQueue.push(message);
            }

            return;
        }

        var children =
            node.querySelectorAll(
                '[data-a-target="chat-line-message"]'
            );

        for (var i = 0; i < children.length; i++) {

            var childMessage =
                parseMessage(children[i]);

            if (childMessage) {
                window.__twitchVoiceQueue.push(
                    childMessage
                );
            }
        }
    }


    function findChat() {

        var selectors = [
            '[data-test-selector="chat-scrollable-area__message-container"]',
            '[data-a-target="chat-scroller"]',
            '[role="log"]'
        ];

        for (var i = 0; i < selectors.length; i++) {

            var element =
                document.querySelector(
                    selectors[i]
                );

            if (element) {
                return element;
            }
        }

        return null;
    }


    function installObserver() {

        var chat = findChat();

        if (!chat) {

            setTimeout(
                installObserver,
                1000
            );

            return;
        }

        if (chat.__twitchVoiceObserver) {
            return;
        }

        chat.__twitchVoiceObserver = true;

        var observer =
            new MutationObserver(
                function (mutations) {

                    for (
                        var i = 0;
                        i < mutations.length;
                        i++
                    ) {

                        var mutation =
                            mutations[i];

                        if (
                            mutation.type !==
                            "childList"
                        ) {
                            continue;
                        }

                        for (
                            var j = 0;
                            j < mutation.addedNodes.length;
                            j++
                        ) {

                            inspectNode(
                                mutation.addedNodes[j]
                            );
                        }
                    }
                }
            );

        observer.observe(
            chat,
            {
                childList: true,
                subtree: true
            }
        );

        console.log(
            "TwitchVoice observer installed"
        );
    }


    installObserver();

})();
"""


# ============================================================
# SIGNALS
# ============================================================

class Signals(QObject):

    message = Signal(str, str)
    status = Signal(str)
    error = Signal(str)
    speaking = Signal(str, bool)


# ============================================================
# TTS
# ============================================================

class TTSEngine:

    def __init__(self, signals):

        self.signals = signals

        self.queue = queue.Queue()

        self.running = True

        self.rate = DEFAULT_RATE

        self.thread = threading.Thread(
            target=self.worker,
            daemon=True
        )

        self.thread.start()


    def set_rate(self, rate):

        self.rate = int(rate)

        print(
            "[TTS] Rate:",
            self.rate
        )


    def find_english_voice(self, sapi):

        voices = sapi.GetVoices()

        for i in range(voices.Count):

            voice = voices.Item(i)

            try:

                name = voice.GetAttribute(
                    "Name"
                )

                language = voice.GetAttribute(
                    "Language"
                )

                name = name.lower()
                language = language.lower()

                if "english" in name:
                    return voice

                if language.startswith("409"):
                    return voice

                if language.startswith("809"):
                    return voice

                if language.startswith("c09"):
                    return voice

                if language.startswith("1009"):
                    return voice

            except Exception:
                pass

        return None


    def speak(self, username, text):

        self.queue.put(
            (
                username,
                text
            )
        )


    def stop(self):

        self.running = False

        self.queue.put(None)


    def worker(self):

        pythoncom.CoInitialize()

        try:

            sapi = win32com.client.Dispatch(
                "SAPI.SpVoice"
            )

            sapi.Volume = 100

            english_voice = (
                self.find_english_voice(
                    sapi
                )
            )

            if english_voice:

                sapi.Voice = english_voice

                try:

                    voice_name = (
                        english_voice.GetAttribute(
                            "Name"
                        )
                    )

                except Exception:

                    voice_name = "English"

            else:

                voice_name = "Windows Default"

            print(
                "[TTS] Voice:",
                voice_name
            )

            self.signals.status.emit(
                "Voice: " + voice_name
            )

            while self.running:

                item = self.queue.get()

                if item is None:
                    break

                username, text = item

                self.signals.speaking.emit(
                    username,
                    True
                )

                try:

                    sapi.Rate = self.rate

                    speech = (
                        username
                        + " says "
                        + text
                    )

                    sapi.Speak(
                        speech
                    )

                except Exception as error:

                    print(
                        "[TTS ERROR]",
                        error
                    )

                finally:

                    self.signals.speaking.emit(
                        username,
                        False
                    )

        finally:

            pythoncom.CoUninitialize()


# ============================================================
# MAIN WINDOW
# ============================================================

class TwitchVoice(QWidget):

    def __init__(self):

        super().__init__()

        self.setWindowTitle(
            "TwitchVoice"
        )

        self.resize(
            1000,
            650
        )

        self.driver = None

        self.running = False

        self.signals = Signals()

        self.signals.message.connect(
            self.on_message
        )

        self.signals.status.connect(
            self.on_status
        )

        self.signals.error.connect(
            self.on_error
        )

        self.signals.speaking.connect(
            self.on_speaking
        )

        self.tts = TTSEngine(
            self.signals
        )

        self.build_ui()

        self.setStyleSheet(
            self.styles()
        )


    # ========================================================
    # STYLE
    # ========================================================

    def styles(self):

        return """
        QWidget {
            background: #09090d;
            color: #eeeeee;
            font-family: "Segoe UI";
        }

        QFrame {
            background: #101016;
            border: 1px solid #24242d;
            border-radius: 12px;
        }

        QLineEdit {
            background: #121219;
            border: 1px solid #292932;
            border-radius: 8px;
            padding: 11px;
            color: #eeeeee;
            font-size: 13px;
        }

        QLineEdit:focus {
            border: 1px solid #8f67dc;
        }

        QPushButton {
            background: #17171f;
            border: 1px solid #30303a;
            border-radius: 8px;
            padding: 11px 22px;
            color: #eeeeee;
        }

        QPushButton:hover {
            background: #22222c;
        }

        QPushButton:disabled {
            color: #55555f;
        }

        QListWidget {
            background: #0d0d12;
            border: none;
            outline: none;
            padding: 8px;
        }

        QListWidget::item {
            padding: 9px;
            border-radius: 6px;
        }

        QListWidget::item:selected {
            background: #1e1e29;
        }

        QSlider::groove:horizontal {
            height: 4px;
            background: #292932;
            border-radius: 2px;
        }

        QSlider::handle:horizontal {
            width: 14px;
            margin: -5px 0;
            border-radius: 7px;
            background: #9a70e8;
        }
        """


    # ========================================================
    # UI
    # ========================================================

    def build_ui(self):

        root = QVBoxLayout(self)

        root.setContentsMargins(
            24,
            20,
            24,
            20
        )

        root.setSpacing(14)


        # HEADER

        header = QHBoxLayout()

        title = QLabel(
            "TWITCHVOICE"
        )

        title.setStyleSheet(
            """
            font-size: 20px;
            font-weight: 700;
            letter-spacing: 4px;
            """
        )

        self.status = QLabel(
            "● OFFLINE"
        )

        self.status.setStyleSheet(
            """
            color: #55555f;
            font-size: 11px;
            letter-spacing: 1px;
            """
        )

        header.addWidget(
            title
        )

        header.addStretch()

        header.addWidget(
            self.status
        )

        root.addLayout(
            header
        )


        # CONNECTION

        connection = QHBoxLayout()

        self.channel = QLineEdit()

        self.channel.setPlaceholderText(
            "Twitch channel name"
        )

        self.connect_button = QPushButton(
            "CONNECT"
        )

        self.connect_button.clicked.connect(
            self.connect_twitch
        )

        connection.addWidget(
            self.channel
        )

        connection.addWidget(
            self.connect_button
        )

        root.addLayout(
            connection
        )


        # SPEED

        speed_panel = QFrame()

        speed_layout = QHBoxLayout(
            speed_panel
        )

        speed_title = QLabel(
            "READOUT SPEED"
        )

        speed_title.setStyleSheet(
            """
            color: #777783;
            font-size: 10px;
            letter-spacing: 2px;
            """
        )

        self.speed = QSlider(
            Qt.Horizontal
        )

        self.speed.setRange(
            MIN_RATE,
            MAX_RATE
        )

        self.speed.setValue(
            DEFAULT_RATE
        )

        self.speed.valueChanged.connect(
            self.speed_changed
        )

        self.speed_value = QLabel(
            "NORMAL"
        )

        self.speed_value.setMinimumWidth(
            80
        )

        self.speed_value.setAlignment(
            Qt.AlignCenter
        )

        speed_layout.addWidget(
            speed_title
        )

        speed_layout.addWidget(
            self.speed
        )

        speed_layout.addWidget(
            self.speed_value
        )

        root.addWidget(
            speed_panel
        )


        # CHAT

        chat_panel = QFrame()

        chat_layout = QVBoxLayout(
            chat_panel
        )

        chat_title = QLabel(
            "LIVE CHAT"
        )

        chat_title.setStyleSheet(
            """
            color: #777783;
            font-size: 10px;
            letter-spacing: 2px;
            """
        )

        self.chat = QListWidget()

        chat_layout.addWidget(
            chat_title
        )

        chat_layout.addWidget(
            self.chat
        )

        root.addWidget(
            chat_panel,
            1
        )


        # FOOTER

        self.footer = QLabel(
            "Ready"
        )

        self.footer.setStyleSheet(
            """
            color: #55555f;
            font-size: 11px;
            """
        )

        root.addWidget(
            self.footer
        )


    # ========================================================
    # SPEED
    # ========================================================

    def speed_changed(self, value):

        self.tts.set_rate(
            value
        )

        if value == 0:

            text = "NORMAL"

        elif value < 0:

            text = "SLOW " + str(value)

        else:

            text = "FAST +" + str(value)

        self.speed_value.setText(
            text
        )


    # ========================================================
    # CONNECT
    # ========================================================

    def connect_twitch(self):

        if self.running:
            return

        channel = (
            self.channel.text().strip()
        )

        channel = channel.replace(
            "#",
            ""
        )

        channel = channel.replace(
            "https://www.twitch.tv/",
            ""
        )

        channel = channel.split(
            "/"
        )[0]

        if not channel:

            self.footer.setText(
                "Enter a Twitch channel."
            )

            return

        self.running = True

        self.connect_button.setEnabled(
            False
        )

        self.status.setText(
            "● CONNECTING"
        )

        self.status.setStyleSheet(
            """
            color: #c89c55;
            font-size: 11px;
            letter-spacing: 1px;
            """
        )

        threading.Thread(
            target=self.twitch_worker,
            args=(channel,),
            daemon=True
        ).start()


    # ========================================================
    # EDGE
    # ========================================================

    def twitch_worker(self, channel):

        try:

            options = Options()

            options.add_argument(
                "--start-maximized"
            )

            options.add_argument(
                "--disable-blink-features=AutomationControlled"
            )

            self.driver = webdriver.Edge(
                options=options
            )

            url = (
                "https://www.twitch.tv/"
                + channel
            )

            self.driver.get(
                url
            )

            self.signals.status.emit(
                "Waiting for Twitch..."
            )

            time.sleep(
                5
            )

            self.driver.execute_script(
                OBSERVER_JS
            )

            self.signals.status.emit(
                "Connected"
            )

            while self.running:

                messages = self.driver.execute_script(
                    """
                    var messages =
                        window.__twitchVoiceQueue
                        || [];

                    window.__twitchVoiceQueue = [];

                    return messages;
                    """
                )

                if messages:

                    for message in messages:

                        username = message.get(
                            "username",
                            ""
                        )

                        text = message.get(
                            "text",
                            ""
                        )

                        if username and text:

                            self.signals.message.emit(
                                username,
                                text
                            )

                time.sleep(
                    POLL_INTERVAL
                )

        except Exception as error:

            self.signals.error.emit(
                str(error)
            )

        finally:

            self.running = False

            try:

                if self.driver:
                    self.driver.quit()

            except Exception:
                pass


    # ========================================================
    # MESSAGE
    # ========================================================

    def on_message(
        self,
        username,
        text
    ):

        self.chat.addItem(
            username + ": " + text
        )

        self.chat.scrollToBottom()

        self.tts.speak(
            username,
            text
        )


    # ========================================================
    # SPEAKING
    # ========================================================

    def on_speaking(
        self,
        username,
        active
    ):

        if active:

            self.footer.setText(
                "Speaking: " + username
            )

            self.status.setText(
                "● SPEAKING"
            )

            self.status.setStyleSheet(
                """
                color: #a678e8;
                font-size: 11px;
                letter-spacing: 1px;
                """
            )

        else:

            self.footer.setText(
                "Ready"
            )

            if self.running:

                self.status.setText(
                    "● LIVE"
                )

                self.status.setStyleSheet(
                    """
                    color: #65c98a;
                    font-size: 11px;
                    letter-spacing: 1px;
                    """
                )


    # ========================================================
    # STATUS
    # ========================================================

    def on_status(self, text):

        self.footer.setText(
            text
        )

        if text == "Connected":

            self.status.setText(
                "● LIVE"
            )

            self.status.setStyleSheet(
                """
                color: #65c98a;
                font-size: 11px;
                letter-spacing: 1px;
                """
            )


    # ========================================================
    # ERROR
    # ========================================================

    def on_error(self, error):

        print()
        print(
            "[TwitchVoice ERROR]"
        )
        print(
            error
        )

        self.status.setText(
            "● ERROR"
        )

        self.status.setStyleSheet(
            """
            color: #d45d6d;
            font-size: 11px;
            letter-spacing: 1px;
            """
        )

        self.footer.setText(
            "Connection error. Check console."
        )

        self.connect_button.setEnabled(
            True
        )


    # ========================================================
    # CLOSE
    # ========================================================

    def closeEvent(self, event):

        self.running = False

        self.tts.stop()

        try:

            if self.driver:
                self.driver.quit()

        except Exception:
            pass

        event.accept()


# ============================================================
# START
# ============================================================

def main():

    app = QApplication(
        sys.argv
    )

    app.setApplicationName(
        "TwitchVoice"
    )

    window = TwitchVoice()

    window.show()

    sys.exit(
        app.exec()
    )


if __name__ == "__main__":
    main()
