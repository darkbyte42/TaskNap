import sys
import os
import platform
import time
import ctypes
from ctypes import wintypes
import winreg
from functools import partial
from datetime import datetime

from PyQt5.QtCore import (
    QTimer,
    QDateTime,
    Qt,
    QSettings,
    QSize,
    QSharedMemory,
    QSystemSemaphore,
)
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QDateTimeEdit,
    QDialog,
    QFormLayout,
    QDialogButtonBox,
    QCheckBox,
    QSpinBox,
    QMessageBox,
    QSystemTrayIcon,
    QMenu,
    QAction,
    QToolBar,
    QGroupBox
)

# -------------------------------------------------------------------------------------
# Windows-specific: System-wide idle detection via GetLastInputInfo
# -------------------------------------------------------------------------------------
class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_uint),
        ("dwTime", ctypes.c_uint),
    ]

def resource_path(relative_path):
    """
    Get the absolute path to the resource.
    Handles both dev mode and PyInstaller-frozen mode.
    """
    if hasattr(sys, '_MEIPASS'):
        # Running in PyInstaller bundle
        base_path = sys._MEIPASS
    else:
        # Running in normal Python
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)

def get_system_idle_time_seconds():
    """
    Returns the number of seconds the system has been idle (on Windows).
    If called on non-Windows, returns 0 just to avoid errors.
    """
    if platform.system().lower().startswith("win"):
        last_input_info = LASTINPUTINFO()
        last_input_info.cbSize = ctypes.sizeof(last_input_info)

        if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(last_input_info)):
            tick_count = ctypes.windll.kernel32.GetTickCount()
            elapsed_ms = tick_count - last_input_info.dwTime
            return elapsed_ms / 1000.0  # convert to seconds
        else:
            return 0
    else:
        # For non-Windows, you'd need a different approach
        return 0

# --------------------------------------------------------------------------
# Custom Date/Time Picker to handle F5 hotkey
# --------------------------------------------------------------------------
class CustomDateTimeEdit(QDateTimeEdit):
    """
    Subclass of QDateTimeEdit that listens for the F5 key.
    When pressed, it updates the widget's date/time to the current system date/time.
    """
    def keyPressEvent(self, event):
        if event.key() == Qt.Key_F5:
            self.setDateTime(QDateTime.currentDateTime())
        else:
            super().keyPressEvent(event)


# --------------------------------------------------------------------------
# Settings Dialog
# --------------------------------------------------------------------------
class SettingsDialog(QDialog):
    """
    A dialog window for user-configurable settings:
      - Info about F5 hotkey
      - Start on system startup (Windows)
      - Enable/disable pre-action notifications and define countdown
      - Enable/disable logging
      - Start minimized to tray
      - Auto-sleep on inactivity
    Settings are saved to an .ini file via QSettings.
    """
    def __init__(self, parent=None, settings=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setFixedSize(QSize(400, 320))

        self.settings = settings
        layout = QFormLayout()

        f5_info_label = QLabel(
            "Press <b>F5</b> in the date/time field to reset to the current local time."
        )
        layout.addRow(f5_info_label)

        self.startup_checkbox = QCheckBox("Start Scheduler at system startup (Windows)")
        layout.addRow(self.startup_checkbox)

        self.pre_action_checkbox = QCheckBox("Enable pre-action notifications")
        layout.addRow(self.pre_action_checkbox)

        self.pre_action_spinbox = QSpinBox()
        self.pre_action_spinbox.setRange(5, 3600)  # seconds: min 5, max 1 hour
        self.pre_action_spinbox.setValue(30)       # default
        self.pre_action_spinbox.setSuffix(" seconds before action")
        layout.addRow("Notification time:", self.pre_action_spinbox)

        self.logging_checkbox = QCheckBox("Enable logging to text file")
        layout.addRow(self.logging_checkbox)

        self.minimize_tray_checkbox = QCheckBox("Start minimized to system tray")
        layout.addRow(self.minimize_tray_checkbox)

        self.auto_sleep_checkbox = QCheckBox("Enable auto-sleep on inactivity")
        layout.addRow(self.auto_sleep_checkbox)

        self.auto_sleep_spinbox = QSpinBox()
        self.auto_sleep_spinbox.setRange(1, 480)   # 1 to 480 minutes (8 hours)
        self.auto_sleep_spinbox.setValue(30)       # default = 30 minutes
        self.auto_sleep_spinbox.setSuffix(" minutes")
        layout.addRow("Inactivity timeout:", self.auto_sleep_spinbox)

        # Dialog buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.on_ok_clicked)
        button_box.rejected.connect(self.reject)
        layout.addRow(button_box)

        self.setLayout(layout)
        self.load_settings_from_ini()

        # On non-Windows, disable the startup checkbox
        if not platform.system().lower().startswith('win'):
            self.startup_checkbox.setEnabled(False)
            self.startup_checkbox.setToolTip("This option is only available on Windows.")

    def load_settings_from_ini(self):
        """
        Load user settings from the QSettings (.ini).
        """
        if not self.settings:
            return

        # Pre-action notifications
        enable_pre_notification = self.settings.value("notifications/enable", False, type=bool)
        self.pre_action_checkbox.setChecked(enable_pre_notification)
        notify_seconds = self.settings.value("notifications/seconds_before", 30, type=int)
        self.pre_action_spinbox.setValue(notify_seconds)

        # Logging
        enable_logging = self.settings.value("logging/enable", False, type=bool)
        self.logging_checkbox.setChecked(enable_logging)

        # Startup
        want_startup = self.settings.value("startup/enable", False, type=bool)
        self.startup_checkbox.setChecked(want_startup)
        if platform.system().lower().startswith('win'):
            is_in_registry = self.is_app_in_startup()
            self.startup_checkbox.setChecked(is_in_registry)

        # Minimize to tray
        minimize_to_tray = self.settings.value("ui/minimize_to_tray", False, type=bool)
        self.minimize_tray_checkbox.setChecked(minimize_to_tray)

        # Auto Sleep on Inactivity
        auto_sleep_enabled = self.settings.value("auto_sleep/enable", False, type=bool)
        self.auto_sleep_checkbox.setChecked(auto_sleep_enabled)

        auto_sleep_minutes = self.settings.value("auto_sleep/timeout_minutes", 30, type=int)
        self.auto_sleep_spinbox.setValue(auto_sleep_minutes)

    def on_ok_clicked(self):
        """
        Apply settings changes (write to .ini and handle registry).
        """
        if not self.settings:
            return

        # Pre-action notifications
        self.settings.setValue("notifications/enable", self.pre_action_checkbox.isChecked())
        self.settings.setValue("notifications/seconds_before", self.pre_action_spinbox.value())

        # Logging
        self.settings.setValue("logging/enable", self.logging_checkbox.isChecked())

        # Startup (Windows registry)
        if platform.system().lower().startswith('win'):
            self.set_app_in_startup(self.startup_checkbox.isChecked())
            self.settings.setValue("startup/enable", self.startup_checkbox.isChecked())

        # Minimize to tray
        self.settings.setValue("ui/minimize_to_tray", self.minimize_tray_checkbox.isChecked())

        # Auto Sleep
        self.settings.setValue("auto_sleep/enable", self.auto_sleep_checkbox.isChecked())
        self.settings.setValue("auto_sleep/timeout_minutes", self.auto_sleep_spinbox.value())

        self.accept()

    # --------------------------------------------------------------------------
    # Registry Manipulation (Windows Only)
    # --------------------------------------------------------------------------
    @staticmethod
    def is_app_in_startup():
        """
        Check if our app is currently set to run at startup in Windows registry.
        Returns True/False.
        """
        run_key = r"Software\Microsoft\Windows\CurrentVersion\Run"
        app_name = "MySchedulerApp"

        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key, 0, winreg.KEY_READ) as key:
                index = 0
                while True:
                    value_name, value_data, value_type = winreg.EnumValue(key, index)
                    if value_name == app_name:
                        return True
                    index += 1
        except WindowsError:
            pass

        return False

    @staticmethod
    def set_app_in_startup(enable):
        """
        Add or remove our app from Windows startup.
        :param enable: bool, True to add to startup, False to remove.
        """
        run_key = r"Software\Microsoft\Windows\CurrentVersion\Run"
        app_name = "MySchedulerApp"

        # Path to Python interpreter or packaged EXE if using PyInstaller
        app_path = sys.executable

        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key, 0, winreg.KEY_WRITE) as key:
                if enable:
                    winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, app_path)
                else:
                    winreg.DeleteValue(key, app_name)
        except WindowsError:
            pass


# --------------------------------------------------------------------------
# Pre-Action Notification Dialog
# --------------------------------------------------------------------------
class PreActionDialog(QMessageBox):
    """
    A simple timed notification dialog that automatically closes after a few seconds.
    Users can click "Cancel" to abort the event, or just ignore it.
    """
    def __init__(self, parent=None, action_type="action", timeout=30):
        super().__init__(parent)
        self.action_type = action_type
        self.timeout = timeout
        self.setWindowTitle("Scheduled Action Notification")
        self.setText(
            f"The system will perform '{action_type}' in {timeout} seconds.\n"
            "Click 'Cancel this Action' to stop it, or wait for it to proceed."
        )
        self.setIcon(QMessageBox.Warning)

        # Add a cancel button
        self.cancel_button = self.addButton("Cancel this Action", QMessageBox.RejectRole)

        # Start a timer to count down & automatically close
        self.close_timer = QTimer(self)
        self.close_timer.timeout.connect(self.on_timeout)
        self.close_timer.start(1000)  # every 1 second

        self.remaining_time = timeout
        self.user_canceled = False  # track if user canceled

    def on_timeout(self):
        """
        Decrease the remaining_time and update text. If it hits 0, we accept.
        """
        self.remaining_time -= 1
        if self.remaining_time <= 0:
            self.close_timer.stop()
            self.accept()  # Did not cancel
        else:
            self.setText(
                f"The system will perform '{self.action_type}' in {self.remaining_time} seconds.\n"
                "Click 'Cancel this Action' to stop it, or wait for it to proceed."
            )

    def reject(self):
        """
        Called when user presses the "Cancel this Action" button.
        """
        self.user_canceled = True
        super().reject()


# --------------------------------------------------------------------------
# Main Scheduler Application (QMainWindow)
# --------------------------------------------------------------------------
class SchedulerApp(QMainWindow):
    """
    A PyQt5 application to schedule system power actions:
      - shutdown
      - restart
      - sleep
    Features:
      - Single-instance check
      - F5 shortcut to reset QDateTimeEdit to current time
      - Pre-action notifications (optional)
      - Logging (optional)
      - System tray icon
      - Config persistence in .ini
      - Balloon notifications on Windows tray
      - Optionally start minimized to tray
      - Auto sleep on inactivity (off by default, now system-wide)
    """
    def __init__(self):
        super().__init__()

        self.settings = QSettings("config.ini", QSettings.IniFormat)

        self.scheduled_events = {}  
        self.last_timer_id = 0

        self.inactivity_timer = QTimer(self)
        self.inactivity_timer.timeout.connect(self.check_inactivity)
        self.inactivity_timer.start(30 * 1000)

        self.setWindowTitle("TaskNap a System Scheduler")
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, "tasknap.ico")
        self.setWindowIcon(QIcon(icon_path))

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)

        # -- Date/Time group box --
        dt_group = QGroupBox("Schedule Time")
        dt_layout = QVBoxLayout(dt_group)
        self.datetime_picker = CustomDateTimeEdit()
        self.datetime_picker.setCalendarPopup(True)
        self.datetime_picker.setDateTime(QDateTime.currentDateTime())
        dt_layout.addWidget(self.datetime_picker)
        main_layout.addWidget(dt_group)

        # -- Info label (for status messages) --
        self.info_label = QLabel("No scheduled events.")
        main_layout.addWidget(self.info_label)

        # -- Actions group box --
        buttons_group = QGroupBox("Actions")
        buttons_layout = QHBoxLayout(buttons_group)

        shutdown_btn = QPushButton("Schedule Shutdown")
        shutdown_btn.clicked.connect(partial(self.schedule_action, 'shutdown'))
        buttons_layout.addWidget(shutdown_btn)

        restart_btn = QPushButton("Schedule Restart")
        restart_btn.clicked.connect(partial(self.schedule_action, 'restart'))
        buttons_layout.addWidget(restart_btn)

        sleep_btn = QPushButton("Schedule Sleep")
        sleep_btn.clicked.connect(partial(self.schedule_action, 'sleep'))
        buttons_layout.addWidget(sleep_btn)

        main_layout.addWidget(buttons_group)

        # -- Cancel & Settings Buttons --
        extra_group = QGroupBox("Manage Schedules")
        extra_layout = QHBoxLayout(extra_group)

        cancel_all_btn = QPushButton("Cancel All Events")
        cancel_all_btn.clicked.connect(self.cancel_all_scheduled_events)
        extra_layout.addWidget(cancel_all_btn)

        settings_btn = QPushButton("Settings")
        settings_btn.clicked.connect(self.open_settings_dialog)
        extra_layout.addWidget(settings_btn)

        main_layout.addWidget(extra_group)

        self.toolbar = QToolBar("Main Toolbar", self)
        self.addToolBar(self.toolbar)

        action_show = QAction(QIcon(icon_path), "Show Window", self)
        action_show.triggered.connect(self.showNormal)
        self.toolbar.addAction(action_show)

        action_settings = QAction("Settings", self)
        action_settings.triggered.connect(self.open_settings_dialog)
        self.toolbar.addAction(action_settings)

        action_cancel = QAction("Cancel All", self)
        action_cancel.triggered.connect(self.cancel_all_scheduled_events)
        self.toolbar.addAction(action_cancel)

        self.init_tray_icon()

        self.load_settings()
        self.resize(400, 300)

    # ----------------------------------------------------------------------
    # System-wide inactivity check
    # ----------------------------------------------------------------------
    def check_inactivity(self):
        """Called periodically to see if we should auto-sleep based on system-wide idle time."""
        auto_sleep_enabled = self.settings.value("auto_sleep/enable", False, type=bool)
        if not auto_sleep_enabled:
            return

        timeout_minutes = self.settings.value("auto_sleep/timeout_minutes", 30, type=int)
        inactivity_threshold = timeout_minutes * 60.0  # convert to seconds

        idle_seconds = get_system_idle_time_seconds()
        # If system-wide idle time >= threshold, trigger sleep
        if idle_seconds >= inactivity_threshold:
            self.perform_system_action("sleep")

    def load_settings(self):
        """
        Load global settings that we may need to apply on startup.
        For instance, if the user wants to start minimized to tray.
        """
        minimize_to_tray = self.settings.value("ui/minimize_to_tray", False, type=bool)
        if minimize_to_tray:
            self.hide()

    # ----------------------------------------------------------------------
    # System Tray Setup
    # ----------------------------------------------------------------------
    def init_tray_icon(self):
        """
        Initialize the system tray icon and its menu.
        """
        self.tray_icon = QSystemTrayIcon(self)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, "tasknap.ico")
        self.tray_icon.setIcon(QIcon(icon_path))

        tray_menu = QMenu()

        show_action = QAction("Show Scheduler", self)
        show_action.triggered.connect(self.show_app)
        tray_menu.addAction(show_action)

        cancel_events_action = QAction("Cancel All Events", self)
        cancel_events_action.triggered.connect(self.cancel_all_scheduled_events)
        tray_menu.addAction(cancel_events_action)

        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(self.quit_app)
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

        # If the user single-clicks the tray icon, show the app
        self.tray_icon.activated.connect(self.on_tray_activated)

    def on_tray_activated(self, reason):
        """
        Handle system tray icon clicks.
        """
        if reason == QSystemTrayIcon.Trigger:  # simple click
            self.show_app()

    def show_app(self):
        """
        Restore/show the application window from the tray.
        """
        self.showNormal()
        self.activateWindow()

    def closeEvent(self, event):
        """
        Override the close event to minimize to tray rather than quit (if desired).
        """
        self.hide()
        event.ignore()  # don't quit, just hide the window

    def quit_app(self):
        """
        Quit the application entirely.
        """
        QApplication.quit()

    # ----------------------------------------------------------------------
    # Settings Dialog
    # ----------------------------------------------------------------------
    def open_settings_dialog(self):
        """
        Open the settings dialog.
        """
        dialog = SettingsDialog(self, self.settings)
        dialog.exec_()

    # ----------------------------------------------------------------------
    # Scheduling Logic
    # ----------------------------------------------------------------------
    def schedule_action(self, action_type):
        """
        Calculates the delay and sets up a QTimer to execute the chosen action.
        """
        scheduled_datetime = self.datetime_picker.dateTime().toPyDateTime()
        current_time = QDateTime.currentDateTime().toPyDateTime()

        delay_seconds = (scheduled_datetime - current_time).total_seconds()
        if delay_seconds <= 0:
            self.info_label.setText('The selected time is in the past. Please choose a future time.')
            return

        self.last_timer_id += 1
        timer_id = self.last_timer_id

        timer = QTimer(self)
        timer.setSingleShot(True)
        timer.timeout.connect(lambda: self.prepare_for_action(timer_id))
        timer.start(int(delay_seconds * 1000))

        self.scheduled_events[timer_id] = (timer, action_type, scheduled_datetime)
        self.log_event(f"Scheduled {action_type} at {scheduled_datetime}")

        msg = (f'Scheduled {action_type} at {scheduled_datetime.strftime("%Y-%m-%d %H:%M:%S")}. '
               f'({len(self.scheduled_events)} events scheduled.)')
        self.info_label.setText(msg)

        # Tray balloon notification
        self.tray_icon.showMessage(
            "Action Scheduled",
            msg,
            QSystemTrayIcon.Information,
            3000
        )

    def prepare_for_action(self, timer_id):
        """
        Called right before we actually perform the action.
        If pre-action notifications are enabled, show the timed dialog.
        """
        if timer_id not in self.scheduled_events:
            return

        _, action_type, scheduled_datetime = self.scheduled_events[timer_id]

        notify_enabled = self.settings.value("notifications/enable", False, type=bool)
        if not notify_enabled:
            self.execute_action(timer_id)
            return

        notify_seconds = self.settings.value("notifications/seconds_before", 30, type=int)

        self.tray_icon.showMessage(
            "Action About to Happen",
            f"{action_type.title()} will occur in {notify_seconds} seconds...",
            QSystemTrayIcon.Warning,
            5000
        )

        dialog = PreActionDialog(self, action_type=action_type, timeout=notify_seconds)
        result = dialog.exec_()

        if dialog.user_canceled:
            self.cancel_event(timer_id, user_triggered=True)
        else:
            self.execute_action(timer_id)

    def execute_action(self, timer_id):
        """
        Performs the system command based on the stored action type,
        then removes the corresponding scheduled event.
        """
        if timer_id not in self.scheduled_events:
            return

        timer, action_type, scheduled_datetime = self.scheduled_events[timer_id]
        del self.scheduled_events[timer_id]

        self.perform_system_action(action_type)
        self.log_event(f"Executed {action_type} at {datetime.now()}")

        remaining = len(self.scheduled_events)
        if remaining > 0:
            msg = (f'Executed {action_type} at {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}. '
                   f'{remaining} event(s) remain scheduled.')
        else:
            msg = (f'Executed {action_type} at {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}. '
                   'No more scheduled events.')
        self.info_label.setText(msg)

        self.tray_icon.showMessage(
            f"Action Executed: {action_type.title()}",
            msg,
            QSystemTrayIcon.Information,
            3000
        )

    def perform_system_action(self, action_type):
        """
        Performs the actual system action (shutdown, restart, sleep),
        handling different OSes if needed.
        """
        os_type = platform.system().lower()

        if action_type == 'shutdown':
            if 'windows' in os_type:
                os.system('shutdown /s /f /t 0')
            else:
                os.system('shutdown -h now')

        elif action_type == 'restart':
            if 'windows' in os_type:
                os.system('shutdown /r /f /t 0')
            else:
                os.system('shutdown -r now')

        elif action_type == 'sleep':
            if 'windows' in os_type:
                # Force system sleep
                os.system('rundll32.exe powrprof.dll,SetSuspendState 0,1,0')
            else:
                os.system('systemctl suspend')

    def cancel_event(self, timer_id, user_triggered=False):
        """
        Cancels a single scheduled event by stopping its timer and removing it.
        """
        if timer_id in self.scheduled_events:
            timer, action_type, scheduled_datetime = self.scheduled_events[timer_id]
            if timer.isActive():
                timer.stop()
            del self.scheduled_events[timer_id]

            msg = f"Canceled {action_type} scheduled at {scheduled_datetime}"
            if user_triggered:
                msg += " (by user)."
            self.info_label.setText(msg)
            self.log_event(msg)

            self.tray_icon.showMessage(
                "Action Canceled",
                msg,
                QSystemTrayIcon.Information,
                3000
            )

    def cancel_all_scheduled_events(self):
        """
        Cancels (stops) all active timers and clears the dictionary of scheduled events.
        """
        if not self.scheduled_events:
            self.info_label.setText('No events to cancel.')
            return

        for timer_id, (timer, action_type, scheduled_datetime) in list(self.scheduled_events.items()):
            if timer.isActive():
                timer.stop()
            del self.scheduled_events[timer_id]
            self.log_event(f"Canceled {action_type} scheduled at {scheduled_datetime}")

        self.info_label.setText('All scheduled events have been canceled.')

        self.tray_icon.showMessage(
            "All Canceled",
            "All scheduled events have been canceled.",
            QSystemTrayIcon.Information,
            3000
        )

    # ----------------------------------------------------------------------
    # Logging
    # ----------------------------------------------------------------------
    def log_event(self, message):
        """
        Append a line to the log file if logging is enabled in settings.
        The log is saved in the same directory as the script, named 'scheduler_log.txt'.
        """
        enable_logging = self.settings.value("logging/enable", False, type=bool)
        if not enable_logging:
            return

        log_path = os.path.join(os.path.dirname(__file__), "scheduler_log.txt")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {message}\n")


# --------------------------------------------------------------------------
# Entry Point
# --------------------------------------------------------------------------
if __name__ == '__main__':
    # Create the QApplication first, so that QMessageBox can appear if needed
    app = QApplication(sys.argv)

    # ------------------------------------------------------------------
    # SINGLE-INSTANCE CHECK
    # ------------------------------------------------------------------
    semaphore_key = "MySchedulerAppSemaphore"
    shared_mem_key = "MySchedulerAppSharedMemory"

    semaphore = QSystemSemaphore(semaphore_key, 1)
    semaphore.acquire()

    shared_memory = QSharedMemory(shared_mem_key)
    is_running = False

    if not shared_memory.create(1):
        is_running = True

    semaphore.release()

    if is_running:
        QMessageBox.information(
            None,
            "Already Running",
            "Another instance of this app is already running.\nExiting now."
        )
        sys.exit(0)

    # no instance is running, proceed
    scheduler = SchedulerApp()

    # show main window unless "start minimized to tray" is active
    minimize_to_tray = scheduler.settings.value("ui/minimize_to_tray", False, type=bool)
    if not minimize_to_tray:
        scheduler.show()

    sys.exit(app.exec_())