from __future__ import annotations

import csv
import json
import sys
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QVBoxLayout,
    QWidget,
)

import azure.cognitiveservices.speech as speechsdk

try:
    import winsound
except ImportError:
    winsound = None

DEFAULT_STYLES = ["Default", "Chat", "Narration"]

LANGUAGE_NAMES = {
    "af": "Afrikaans",
    "am": "Amharic",
    "ar": "Arabic",
    "az": "Azerbaijani",
    "bg": "Bulgarian",
    "bn": "Bengali",
    "bs": "Bosnian",
    "ca": "Catalan",
    "cs": "Czech",
    "cy": "Welsh",
    "da": "Danish",
    "de": "German",
    "el": "Greek",
    "en": "English",
    "es": "Spanish",
    "et": "Estonian",
    "fa": "Persian",
    "fi": "Finnish",
    "fil": "Filipino",
    "fr": "French",
    "ga": "Irish",
    "gl": "Galician",
    "gu": "Gujarati",
    "he": "Hebrew",
    "hi": "Hindi",
    "hr": "Croatian",
    "hu": "Hungarian",
    "id": "Indonesian",
    "is": "Icelandic",
    "it": "Italian",
    "ja": "Japanese",
    "jv": "Javanese",
    "ka": "Georgian",
    "kk": "Kazakh",
    "km": "Khmer",
    "kn": "Kannada",
    "ko": "Korean",
    "lo": "Lao",
    "lt": "Lithuanian",
    "lv": "Latvian",
    "mk": "Macedonian",
    "ml": "Malayalam",
    "mn": "Mongolian",
    "mr": "Marathi",
    "ms": "Malay",
    "my": "Burmese",
    "nb": "Norwegian Bokmål",
    "ne": "Nepali",
    "nl": "Dutch",
    "pl": "Polish",
    "ps": "Pashto",
    "pt": "Portuguese",
    "ro": "Romanian",
    "ru": "Russian",
    "si": "Sinhala",
    "sk": "Slovak",
    "sl": "Slovenian",
    "so": "Somali",
    "sq": "Albanian",
    "sr": "Serbian",
    "su": "Sundanese",
    "sv": "Swedish",
    "sw": "Swahili",
    "ta": "Tamil",
    "te": "Telugu",
    "th": "Thai",
    "tr": "Turkish",
    "uk": "Ukrainian",
    "ur": "Urdu",
    "uz": "Uzbek",
    "vi": "Vietnamese",
    "zh": "Chinese",
    "zu": "Zulu",
}


@dataclass
class CsvRow:
    path: str
    text_to_play: str


class SynthesisWorker(QThread):
    log = pyqtSignal(str)
    done = pyqtSignal()
    failed = pyqtSignal(str)

    def __init__(
        self,
        app_root: Path,
        csv_file_path: Path,
        output_dir: Path,
        key: str,
        region_value: str,
        language_code: str,
        language_name: str,
        selected_region: str,
        short_name: str,
        style: str,
        speed: float,
        post_silence: int,
        pre_silence: int,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.app_root = app_root
        self.csv_file_path = csv_file_path
        self.output_dir = output_dir
        self.key = key
        self.region_value = region_value
        self.language_code = language_code
        self.language_name = language_name
        self.selected_region = selected_region
        self.short_name = short_name
        self.style = style
        self.speed = speed
        self.post_silence = post_silence
        self.pre_silence = pre_silence

    def run(self) -> None:
        try:
            self._run_impl()
            self.done.emit()
        except Exception:
            self.failed.emit(traceback.format_exc())

    def _run_impl(self) -> None:
        base_output = (
            self.output_dir
            / self.language_code
            / self.selected_region
            / self.short_name.split("-")[-1]
        )
        base_output.mkdir(parents=True, exist_ok=True)

        csv_path = self._resolve_csv_path()
        current_rows = self._read_csv_rows(csv_path)

        data_dir = self.app_root / "data"
        data_dir.mkdir(parents=True, exist_ok=True)
        last_csv_path = data_dir / "lastCsvData.csv"
        if not last_csv_path.exists():
            self._write_rows(last_csv_path, current_rows)

        self.log.emit("Checking for changes in .csv file from last run...")
        last_rows = self._read_csv_rows(last_csv_path)
        last_map = {row.path: row.text_to_play for row in last_rows}

        changed_rows: list[CsvRow] = []
        new_rows: list[CsvRow] = []
        for row in current_rows:
            if row.path not in last_map:
                new_rows.append(row)
                self.log.emit(f"{row.path} is a new row.")
                continue
            if last_map[row.path] != row.text_to_play:
                changed_rows.append(row)
                self.log.emit(f"{row.path} text to play has changed. Removing old audio file.")
                old_file = base_output / row.path.replace(":", "_")
                if old_file.exists():
                    old_file.unlink(missing_ok=True)

        if changed_rows and new_rows:
            target_rows = changed_rows + new_rows
            self.log.emit("Changes detected in .csv file. Synthesizing new and changed rows.")
        elif changed_rows:
            target_rows = changed_rows
            self.log.emit("Changes detected in .csv file. Synthesizing changed rows.")
        elif new_rows:
            target_rows = new_rows
            self.log.emit("Changes detected in .csv file. Synthesizing new rows.")
        else:
            target_rows = current_rows
            self.log.emit("No changes detected in .csv file. Only missing files will be synthesized.")

        speech_config = speechsdk.SpeechConfig(subscription=self.key, region=self.region_value)
        speech_config.speech_synthesis_voice_name = self.short_name
        failed_outputs: list[str] = []

        for row in target_rows:
            safe_relative_path = Path(row.path.replace(":", "_"))
            if safe_relative_path.is_absolute() or ".." in safe_relative_path.parts:
                self.log.emit(f"Skipping unsafe CSV path: {row.path}")
                continue
            output_file = base_output / safe_relative_path
            output_file.parent.mkdir(parents=True, exist_ok=True)

            if output_file.exists():
                self.log.emit(f"File {output_file} already exists. Skipping synthesis.")
                continue

            self.log.emit(f"Synthesizing {output_file}...")
            ssml = build_ssml(
                language_code=self.language_code,
                region=self.selected_region,
                short_name=self.short_name,
                text=row.text_to_play,
                style=self.style,
                speed=self.speed,
                pre_silence=self.pre_silence,
                post_silence=self.post_silence,
            )

            retry_count = 0
            while True:
                audio_config = speechsdk.audio.AudioOutputConfig(filename=str(output_file))
                synthesizer = speechsdk.SpeechSynthesizer(
                    speech_config=speech_config,
                    audio_config=audio_config,
                )
                result = synthesizer.speak_ssml_async(ssml).get()
                if result is None:
                    retry_count += 1
                    self.log.emit(f"Synthesis returned no result for {output_file}. Retry count: {retry_count}")
                    if retry_count >= 3:
                        output_file.unlink(missing_ok=True)
                        self.log.emit("Audio synthesis failed. Retry limit reached for this file.")
                        failed_outputs.append(str(output_file))
                        break
                    continue

                if result.reason != speechsdk.ResultReason.SynthesizingAudioCompleted:
                    retry_count += 1
                    self.log.emit(f"Synthesis failed for {output_file}. Retry count: {retry_count}")
                elif not output_file.exists() or output_file.stat().st_size < 1024:
                    retry_count += 1
                    self.log.emit("File generated is invalid. Retrying...")
                else:
                    self.log.emit("Done")
                    break

                if retry_count >= 3:
                    output_file.unlink(missing_ok=True)
                    self.log.emit("Audio synthesis failed. Retry limit reached for this file.")
                    failed_outputs.append(str(output_file))
                    break

        if failed_outputs:
            preview = ", ".join(failed_outputs[:5])
            if len(failed_outputs) > 5:
                preview += ", ..."
            raise RuntimeError(
                f"Synthesis failed for {len(failed_outputs)} file(s) after 3 retries. "
                f"Examples: {preview}"
            )

        self._write_rows(last_csv_path, current_rows)
        settings_payload = {
            "style": self.style,
            "multiplier": self.speed,
            "trailingSilence": self.post_silence,
            "leadingSilence": self.pre_silence,
        }
        settings_path = base_output / "settings.json"
        settings_path.write_text(json.dumps(settings_payload, indent=2), encoding="utf-8")
        self.log.emit(f"Settings saved to {settings_path}")

    def _resolve_csv_path(self) -> Path:
        csv_path = self.csv_file_path
        if not csv_path.exists() or not csv_path.is_file() or csv_path.suffix.lower() != ".csv":
            raise FileNotFoundError(f"CSV file not found or invalid: {csv_path}")
        self.log.emit(f"Using CSV: {csv_path}")
        return csv_path

    def _read_csv_rows(self, csv_path: Path) -> list[CsvRow]:
        rows: list[CsvRow] = []
        with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
            reader = csv.DictReader(file)
            for row in reader:
                path_value = (row.get("path") or "").strip()
                text_value = (row.get("text to play") or "").strip()
                if not path_value:
                    continue
                rows.append(CsvRow(path=path_value, text_to_play=text_value))
        return rows

    def _write_rows(self, csv_path: Path, rows: list[CsvRow]) -> None:
        with csv_path.open("w", encoding="utf-8", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=["path", "text to play"])
            writer.writeheader()
            for row in rows:
                writer.writerow({"path": row.path, "text to play": row.text_to_play})


class PreviewWorker(QThread):
    done = pyqtSignal(str)
    failed = pyqtSignal(str)

    def __init__(
        self,
        key: str,
        region_value: str,
        short_name: str,
        ssml: str,
        sample_wav_path: Path,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.key = key
        self.region_value = region_value
        self.short_name = short_name
        self.ssml = ssml
        self.sample_wav_path = sample_wav_path

    def run(self) -> None:
        try:
            self._run_impl()
        except Exception:
            self.failed.emit(traceback.format_exc())

    def _run_impl(self) -> None:
        speech_config = speechsdk.SpeechConfig(subscription=self.key, region=self.region_value)
        speech_config.speech_synthesis_voice_name = self.short_name

        audio_config = speechsdk.audio.AudioOutputConfig(filename=str(self.sample_wav_path))
        synthesizer = speechsdk.SpeechSynthesizer(
            speech_config=speech_config,
            audio_config=audio_config,
        )
        result = synthesizer.speak_ssml_async(self.ssml).get()
        if result is None:
            raise RuntimeError("Voice preview failed: no result returned from Azure Speech.")
        if result.reason != speechsdk.ResultReason.SynthesizingAudioCompleted:
            raise RuntimeError("Voice preview failed.")
        if not self.sample_wav_path.exists() or self.sample_wav_path.stat().st_size == 0:
            raise RuntimeError("Voice preview failed: no audio was written.")

        if winsound is not None:
            winsound.PlaySound(str(self.sample_wav_path), winsound.SND_FILENAME | winsound.SND_ASYNC)
            self.done.emit("Voice preview played.")
            return

        self.done.emit(f"Voice preview generated at {self.sample_wav_path}")


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Advanced Audio Suite")
        self.resize(880, 620)

        if getattr(sys, "frozen", False):
            self.app_root = Path(sys.executable).resolve().parent
        else:
            self.app_root = Path(__file__).resolve().parent
        self.config_dir = self.app_root / "config"
        self.data_dir = self.app_root / "data"
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.data_dir.mkdir(parents=True, exist_ok=True)

        self.credentials_path = self.config_dir / "credentials.json"
        self.config_path = self.config_dir / "config.json"
        self.voices_path = self.data_dir / "voices.json"

        self.voices: list[dict[str, Any]] = []
        self.language_map: dict[str, str] = {}
        self.worker: SynthesisWorker | None = None
        self.preview_worker: PreviewWorker | None = None

        self._build_ui()
        self.key, self.azure_region = self._ensure_credentials()
        self._load_or_fetch_voices()
        self._populate_languages()
        self._load_config_defaults()

    def _build_ui(self) -> None:
        central = QWidget(self)
        self.setCentralWidget(central)

        layout = QVBoxLayout(central)
        title = QLabel("Advanced Audio Suite")
        title.setStyleSheet("font-size: 22px; font-weight: 600;")
        layout.addWidget(title)

        grid = QGridLayout()
        form = QFormLayout()

        self.cmb_language = QComboBox()
        self.cmb_region = QComboBox()
        self.cmb_voice = QComboBox()
        self.cmb_style = QComboBox()
        self.txt_speed = QLineEdit()
        self.txt_post_silence = QLineEdit()
        self.txt_pre_silence = QLineEdit()
        self.txt_sample_text = QLineEdit()
        self.txt_csv_file = QLineEdit()
        self.txt_output_dir = QLineEdit()
        self.btn_browse_csv = QPushButton("Browse...")
        self.btn_browse_output = QPushButton("Browse...")

        form.addRow("Language:", self.cmb_language)
        form.addRow("Region:", self.cmb_region)
        form.addRow("Voice:", self.cmb_voice)
        form.addRow("Voice Style:", self.cmb_style)
        form.addRow("Speed Multiplier:", self.txt_speed)
        form.addRow("Trailing Silence (ms):", self.txt_post_silence)
        form.addRow("Leading Silence (ms):", self.txt_pre_silence)

        csv_row = QWidget()
        csv_layout = QHBoxLayout(csv_row)
        csv_layout.setContentsMargins(0, 0, 0, 0)
        csv_layout.addWidget(self.txt_csv_file)
        csv_layout.addWidget(self.btn_browse_csv)
        form.addRow("CSV File:", csv_row)

        output_row = QWidget()
        output_layout = QHBoxLayout(output_row)
        output_layout.setContentsMargins(0, 0, 0, 0)
        output_layout.addWidget(self.txt_output_dir)
        output_layout.addWidget(self.btn_browse_output)
        form.addRow("Output Directory:", output_row)

        grid.addLayout(form, 0, 0)
        layout.addLayout(grid)

        button_layout = QHBoxLayout()
        self.btn_play = QPushButton("Preview Selected Voice")
        self.btn_start = QPushButton("Start Synthesis")
        self.txt_sample_text.setPlaceholderText("Enter sample text...")
        self.txt_sample_text.setMinimumWidth(280)
        button_layout.addWidget(self.btn_play)
        button_layout.addWidget(self.txt_sample_text)
        button_layout.addStretch(1)
        button_layout.addWidget(self.btn_start)
        layout.addLayout(button_layout)

        self.output = QPlainTextEdit()
        self.output.setReadOnly(True)
        layout.addWidget(self.output)

        self.cmb_style.addItems(DEFAULT_STYLES)

        self.cmb_language.currentTextChanged.connect(self._on_language_changed)
        self.cmb_region.currentTextChanged.connect(self._on_region_changed)
        self.cmb_voice.currentTextChanged.connect(self._on_voice_changed)
        self.btn_browse_csv.clicked.connect(self._choose_csv_file)
        self.btn_browse_output.clicked.connect(self._choose_output_directory)
        self.btn_play.clicked.connect(self._on_play_sample)
        self.btn_start.clicked.connect(self._on_start_synthesis)

    def _log(self, message: str) -> None:
        self.output.appendPlainText(message)

    def _ensure_credentials(self) -> tuple[str, str]:
        default_creds = {"Key": "yourkey", "Region": "yourregion"}
        if not self.credentials_path.exists():
            self.credentials_path.write_text(json.dumps(default_creds, indent=2), encoding="utf-8")

        creds = json.loads(self.credentials_path.read_text(encoding="utf-8"))
        key = (creds.get("Key") or "").strip()
        region_value = (creds.get("Region") or "").strip()

        if not key or not region_value or key == "yourkey" or region_value == "yourregion":
            key, ok = QInputDialog.getText(
                self,
                "Azure Key",
                "Enter your Azure Speech key:",
            )
            if not ok or not key.strip():
                raise RuntimeError("Azure Speech key is required.")

            region_value, ok = QInputDialog.getText(
                self,
                "Azure Region",
                "Enter your Azure Speech region:",
            )
            if not ok or not region_value.strip():
                raise RuntimeError("Azure Speech region is required.")

            payload = {"Key": key.strip(), "Region": region_value.strip()}
            self.credentials_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
            key = payload["Key"]
            region_value = payload["Region"]

        return key, region_value

    def _load_or_fetch_voices(self) -> None:
        if self.voices_path.exists():
            self.voices = json.loads(self.voices_path.read_text(encoding="utf-8"))
            self._log("Reading voices.json... OK")
            return

        self._log("Getting voices from Azure Speech SDK...")
        speech_config = speechsdk.SpeechConfig(subscription=self.key, region=self.azure_region)
        synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=None)
        result = synthesizer.get_voices_async("").get()
        if result is None:
            raise RuntimeError("Failed to retrieve voices from Azure Speech SDK.")
        if result.reason != speechsdk.ResultReason.VoicesListRetrieved:
            raise RuntimeError("Failed to retrieve voices from Azure Speech SDK.")

        serialized: list[dict[str, Any]] = []
        for voice in result.voices:
            locale = str(getattr(voice, "locale", "") or "")
            locale_name = str(
                getattr(voice, "locale_name", None)
                or getattr(voice, "local_name", None)
                or locale
            )
            short_name = str(
                getattr(voice, "short_name", None)
                or getattr(voice, "name", None)
                or ""
            )
            style_list_value = getattr(voice, "style_list", None) or []
            serialized.append(
                {
                    "Locale": locale,
                    "LocaleName": locale_name,
                    "ShortName": short_name,
                    "StyleList": list(style_list_value),
                }
            )

        self.voices = serialized
        self.voices_path.write_text(json.dumps(serialized, indent=2), encoding="utf-8")
        self._log("Saved voices to data/voices.json")

    def _populate_languages(self) -> None:
        language_map: dict[str, str] = {}
        for voice in self.voices:
            locale = str(voice.get("Locale", "") or "")
            if "-" not in locale:
                continue
            language_code = locale.split("-")[0]
            if language_code not in language_map:
                language_map[language_code] = language_name_from_code(language_code)

        self.language_map = language_map
        languages = sorted(set(language_map.values()))
        self.cmb_language.clear()
        self.cmb_language.addItems(languages)

    def _load_config_defaults(self) -> None:
        default_csv_file = self._default_csv_file_path()
        defaults = {
            "Language": "English",
            "Region": "AU",
            "Voice": "ElsieNeural",
            "Style": "Default",
            "Speed": 1.25,
            "PostSilence": 25,
            "PreSilence": 0,
            "SampleText": "Welcome to Ethos",
            "InputCsvFile": str(default_csv_file) if default_csv_file else "",
            "OutputDirectory": str((self.app_root / "out").resolve()),
        }
        if self.config_path.exists():
            config = defaults | json.loads(self.config_path.read_text(encoding="utf-8"))
        else:
            config = defaults

        if not config.get("InputCsvFile") and config.get("InputDirectory"):
            legacy_dir = Path(str(config.get("InputDirectory")))
            if legacy_dir.exists() and legacy_dir.is_dir():
                legacy_csv_files = sorted(
                    path for path in legacy_dir.iterdir() if path.is_file() and path.suffix.lower() == ".csv"
                )
                if legacy_csv_files:
                    config["InputCsvFile"] = str(legacy_csv_files[0])

        self.cmb_language.setCurrentText(str(config["Language"]))
        self._refresh_regions()
        self.cmb_region.setCurrentText(str(config["Region"]))
        self._refresh_voices()
        self.cmb_voice.setCurrentText(str(config["Voice"]))
        self._refresh_styles()
        self.cmb_style.setCurrentText(str(config["Style"]))

        self.txt_speed.setText(str(config["Speed"]))
        self.txt_post_silence.setText(str(config["PostSilence"]))
        self.txt_pre_silence.setText(str(config["PreSilence"]))
        self.txt_sample_text.setText(str(config["SampleText"]))
        self.txt_csv_file.setText(str(config["InputCsvFile"]))
        self.txt_output_dir.setText(str(config["OutputDirectory"]))

    def _choose_csv_file(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "Select Input CSV File",
            self.txt_csv_file.text().strip() or str(self.app_root),
            "CSV Files (*.csv)",
        )
        if selected:
            self.txt_csv_file.setText(selected)

    def _choose_output_directory(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            "Select Output Directory",
            self.txt_output_dir.text().strip() or str(self.app_root),
        )
        if selected:
            self.txt_output_dir.setText(selected)

    def _selected_language_code(self) -> str:
        selected_language = self.cmb_language.currentText().strip()
        for code, name in self.language_map.items():
            if name == selected_language:
                return code
        return "en"

    def _refresh_regions(self) -> None:
        language_code = self._selected_language_code()
        regions = sorted(
            {
                str(voice.get("Locale", "")).split("-")[1]
                for voice in self.voices
                if str(voice.get("Locale", "")).startswith(f"{language_code}-") and "-" in str(voice.get("Locale", ""))
            }
        )
        current = self.cmb_region.currentText()
        self.cmb_region.blockSignals(True)
        self.cmb_region.clear()
        self.cmb_region.addItems(regions)
        if current in regions:
            self.cmb_region.setCurrentText(current)
        elif regions:
            self.cmb_region.setCurrentIndex(0)
        self.cmb_region.blockSignals(False)

    def _refresh_voices(self) -> None:
        language_code = self._selected_language_code()
        region = self.cmb_region.currentText().strip()
        matches = [
            voice
            for voice in self.voices
            if str(voice.get("Locale", "")).startswith(f"{language_code}-{region}")
        ]
        short_names = [str(voice.get("ShortName", "")).split("-")[-1] for voice in matches if voice.get("ShortName")]

        current = self.cmb_voice.currentText()
        self.cmb_voice.blockSignals(True)
        self.cmb_voice.clear()
        self.cmb_voice.addItems(short_names)
        if current in short_names:
            self.cmb_voice.setCurrentText(current)
        elif short_names:
            self.cmb_voice.setCurrentIndex(0)
        self.cmb_voice.blockSignals(False)

    def _refresh_styles(self) -> None:
        selected_short_name = self._current_short_name()
        style_list: list[str] = []
        if selected_short_name:
            for voice in self.voices:
                if voice.get("ShortName") == selected_short_name:
                    style_list = voice.get("StyleList") or []
                    break

        if not style_list:
            style_options = DEFAULT_STYLES
        else:
            normalized_styles = {
                str(style).strip().title()
                for style in style_list
                if str(style).strip()
            }
            normalized_styles.discard("Default")
            style_options = ["Default"] + sorted(normalized_styles)

        current = self.cmb_style.currentText()
        self.cmb_style.clear()
        self.cmb_style.addItems(style_options)
        if current in style_options:
            self.cmb_style.setCurrentText(current)

    def _on_language_changed(self) -> None:
        self._refresh_regions()
        self._refresh_voices()
        self._refresh_styles()

    def _on_region_changed(self) -> None:
        self._refresh_voices()
        self._refresh_styles()

    def _on_voice_changed(self) -> None:
        self._refresh_styles()

    def _current_short_name(self) -> str:
        language_code = self._selected_language_code()
        region = self.cmb_region.currentText().strip()
        voice_tail = self.cmb_voice.currentText().strip()
        if not language_code or not region or not voice_tail:
            return ""
        return f"{language_code}-{region}-{voice_tail}"

    def _save_config(self) -> None:
        payload = {
            "Language": self.cmb_language.currentText().strip(),
            "Region": self.cmb_region.currentText().strip(),
            "Voice": self.cmb_voice.currentText().strip(),
            "Style": self.cmb_style.currentText().strip(),
            "Speed": float(self.txt_speed.text().strip()),
            "PostSilence": int(self.txt_post_silence.text().strip()),
            "PreSilence": int(self.txt_pre_silence.text().strip()),
            "SampleText": self.txt_sample_text.text().strip(),
            "InputCsvFile": self.txt_csv_file.text().strip(),
            "OutputDirectory": self.txt_output_dir.text().strip(),
        }
        self.config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def _on_play_sample(self) -> None:
        if self.preview_worker and self.preview_worker.isRunning():
            QMessageBox.information(self, "Busy", "Voice preview is already running.")
            return

        try:
            self._save_config()
            speed = float(self.txt_speed.text().strip())
            post_silence = int(self.txt_post_silence.text().strip())
            pre_silence = int(self.txt_pre_silence.text().strip())
        except ValueError:
            QMessageBox.critical(self, "Invalid Input", "Speed and silence values must be numeric.")
            return

        language_code = self._selected_language_code()
        region = self.cmb_region.currentText().strip()
        short_name = self._current_short_name()
        if not short_name:
            QMessageBox.critical(self, "Missing Voice", "Please select a valid voice.")
            return

        text = self.txt_sample_text.text().strip()
        if not text:
            QMessageBox.critical(self, "Sample Text", "Please enter sample text to speak.")
            return

        style = self.cmb_style.currentText().strip() or "Default"
        ssml = build_ssml(
            language_code=language_code,
            region=region,
            short_name=short_name,
            text=text,
            style=style,
            speed=speed,
            pre_silence=pre_silence,
            post_silence=post_silence,
        )

        self.btn_play.setEnabled(False)
        self._log("Generating voice preview...")
        self.preview_worker = PreviewWorker(
            key=self.key,
            region_value=self.azure_region,
            short_name=short_name,
            ssml=ssml,
            sample_wav_path=self.data_dir / "sample_preview.wav",
            parent=self,
        )
        self.preview_worker.done.connect(self._on_preview_done)
        self.preview_worker.failed.connect(self._on_preview_failed)
        self.preview_worker.start()

    def _on_preview_done(self, message: str) -> None:
        self.btn_play.setEnabled(True)
        self._log(message)

    def _on_preview_failed(self, details: str) -> None:
        self.btn_play.setEnabled(True)
        QMessageBox.critical(self, "Speech Error", details)

    def _on_start_synthesis(self) -> None:
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, "Busy", "Synthesis is already running.")
            return

        try:
            self._save_config()
            speed = float(self.txt_speed.text().strip())
            post_silence = int(self.txt_post_silence.text().strip())
            pre_silence = int(self.txt_pre_silence.text().strip())
        except ValueError:
            QMessageBox.critical(self, "Invalid Input", "Speed and silence values must be numeric.")
            return

        language_code = self._selected_language_code()
        language_name = self.cmb_language.currentText().strip()
        region = self.cmb_region.currentText().strip()
        short_name = self._current_short_name()
        style = self.cmb_style.currentText().strip() or "Default"
        csv_file_path = Path(self.txt_csv_file.text().strip())
        output_dir = Path(self.txt_output_dir.text().strip())

        if not csv_file_path.exists() or not csv_file_path.is_file() or csv_file_path.suffix.lower() != ".csv":
            QMessageBox.critical(self, "CSV File", "Please select a valid CSV file.")
            return

        output_dir.mkdir(parents=True, exist_ok=True)

        if not short_name:
            QMessageBox.critical(self, "Missing Voice", "Please select a valid voice.")
            return

        self.btn_start.setEnabled(False)
        self._log("Starting synthesis...")
        self.worker = SynthesisWorker(
            app_root=self.app_root,
            csv_file_path=csv_file_path,
            output_dir=output_dir,
            key=self.key,
            region_value=self.azure_region,
            language_code=language_code,
            language_name=language_name,
            selected_region=region,
            short_name=short_name,
            style=style,
            speed=speed,
            post_silence=post_silence,
            pre_silence=pre_silence,
            parent=self,
        )
        self.worker.log.connect(self._log)
        self.worker.done.connect(self._on_synthesis_done)
        self.worker.failed.connect(self._on_synthesis_failed)
        self.worker.start()

    def _default_csv_file_path(self) -> Path | None:
        in_dir = self.app_root / "in"
        if not in_dir.exists() or not in_dir.is_dir():
            return None
        csv_files = sorted(path for path in in_dir.iterdir() if path.is_file() and path.suffix.lower() == ".csv")
        if not csv_files:
            return None
        return csv_files[0]

    def _on_synthesis_done(self) -> None:
        self.btn_start.setEnabled(True)
        language_code = self._selected_language_code()
        region = self.cmb_region.currentText().strip()
        output_dir = self.txt_output_dir.text().strip()
        self._log(
            f"Speech synthesis complete. Synthesized audio can be found in "
            f"{output_dir}/{language_code}/{region}."
        )
        self._cleanup_logs()

    def _on_synthesis_failed(self, details: str) -> None:
        self.btn_start.setEnabled(True)
        self._log("Synthesis failed.")
        QMessageBox.critical(self, "Synthesis Error", details)
        self._cleanup_logs()

    def _cleanup_logs(self) -> None:
        for path in self.app_root.glob("log-*"):
            if path.is_file():
                path.unlink(missing_ok=True)
        self._log("Cleaning up...")


def speed_multiplier_to_rate(speed: float) -> str:
    percentage = (speed - 1.0) * 100.0
    return f"{percentage:+.2f}%"


def build_ssml(
    language_code: str,
    region: str,
    short_name: str,
    text: str,
    style: str,
    speed: float,
    pre_silence: int,
    post_silence: int,
) -> str:
    xml_lang = f"{language_code}-{region}"
    rate = speed_multiplier_to_rate(speed)
    if style.lower() == "default":
        style_wrapper_start = ""
        style_wrapper_end = ""
    else:
        style_wrapper_start = f"<mstts:express-as style='{style.lower()}' styledegree='2'>"
        style_wrapper_end = "</mstts:express-as>"

    return (
        "<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' "
        "xmlns:mstts='https://www.w3.org/2001/mstts' "
        f"xml:lang='{xml_lang}'>"
        f"<voice name='{short_name}'>"
        f"{style_wrapper_start}"
        f"<lang xml:lang='{xml_lang}'>"
        f"<prosody rate='{rate}'>"
        f"<mstts:silence type='Leading-exact' value='{pre_silence}ms'/>"
        f"{escape_xml_text(text)}"
        f"<mstts:silence type='Trailing-exact' value='{post_silence}ms'/>"
        "</prosody>"
        "</lang>"
        f"{style_wrapper_end}"
        "</voice>"
        "</speak>"
    )


def escape_xml_text(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def language_name_from_code(language_code: str) -> str:
    code = (language_code or "").strip().lower()
    if not code:
        return "Unknown"
    return LANGUAGE_NAMES.get(code, code)


def main() -> int:
    app = QApplication(sys.argv)
    try:
        window = MainWindow()
    except Exception as exc:
        QMessageBox.critical(None, "Startup Error", str(exc))
        return 1
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
