import obspython as obs
import sys
import win32api
from PIL import ImageGrab
import time
import continuous_threading
import random

version = "2.3"

g = lambda: ...
g.settings = None

monitors = []
monitor = None

slide_scene = ""
active = False

screen_sourcename = ""
camera_sourcename = ""
camera_blur = 25

screen_visible = True
camera_locked = True

previous_image = []
timestamp = 0

slide_visible_duration = 10

refresh_interval = 0.1
periodic_thread = None

fadeout_duration = 0.25
fadeout_timestamp = 0

hotkey = obs.OBS_INVALID_HOTKEY_ID
holding_hotkey = False

def hotkey_callback(pressed):
	global holding_hotkey
	if active:
		holding_hotkey = pressed
		if pressed:
			update_opacity(100)
		else:
			fadeout()

DEFAULT_STATUS = 0
BLACK_STATUS = 1
NEWSLIDE_STATUS = 2
status = DEFAULT_STATUS

def update_opacity(value):
	global screen_visible
	global camera_locked
	global timestamp
	global status

	status = DEFAULT_STATUS

	value = int(value)

	screen_visible = (value == 100)
	if screen_visible:
		obs.timer_remove(fadeout_callback)
	timestamp = time.time() if screen_visible else 0
	camera_locked = (value != 0)

	set_filter_value(screen_sourcename, "Color Correction", "opacity", value)
	set_filter_value(camera_sourcename, "Blur", "Filter.Blur.Size", int(camera_blur * (value / 100)))

def set_filter_value(source_name, filter_name, filter_field_name, value):
	source = obs.obs_get_source_by_name(source_name)
	if source is not None:
		filter = obs.obs_source_get_filter_by_name(source, filter_name)
		if filter is not None:
			# Get the settings data object for the filter
			filter_settings = obs.obs_source_get_settings(filter)

			# Update the hue_shift property and update the filter with the new settings
			obs.obs_data_set_int(filter_settings, filter_field_name, value)
			obs.obs_source_update(filter, filter_settings)

			# Release the resources
			obs.obs_data_release(filter_settings)
			obs.obs_source_release(filter)
		obs.obs_source_release(source)

def get_filter_value(source_name, filter_name, filter_field_name):
	value = None

	source = obs.obs_get_source_by_name(source_name)
	if source is not None:
		filter = obs.obs_source_get_filter_by_name(source, filter_name)
		if filter is not None:
			# Get the settings data object for the filter
			filter_settings = obs.obs_source_get_settings(filter)

			# Update the hue_shift property and update the filter with the new settings
			value = obs.obs_data_get_int(filter_settings, filter_field_name)

			# Release the resources
			obs.obs_data_release(filter_settings)
			obs.obs_source_release(filter)
		obs.obs_source_release(source)
	return value

def fadeout_callback():
	if time.time() <= fadeout_timestamp + fadeout_duration:
		update_opacity(((time.time() - fadeout_timestamp) / fadeout_duration * -100) + 100)
	else:
		update_opacity(0)
		obs.timer_remove(fadeout_callback)

def fadeout():
	global fadeout_timestamp
	global status
	status = DEFAULT_STATUS

	obs.timer_remove(fadeout_callback)
	fadeout_timestamp = time.time()
	obs.timer_add(fadeout_callback, 10)

def update_backend():
	global previous_image
	global status

	im = ImageGrab.grab(monitors[monitor].pyRect, False, True, None)

	if im != previous_image:
		previous_image = im
		if im.convert("L").getextrema() == (0,0):
			status = BLACK_STATUS
		else:
			status = NEWSLIDE_STATUS

def update_ui():

	if status == BLACK_STATUS:
		update_opacity(0)
	elif status == NEWSLIDE_STATUS:
		update_opacity(100)
	elif time.time() >= timestamp + slide_visible_duration and screen_visible and not holding_hotkey:
		fadeout()

def activate_timer():
	global active
	global previous_image
	global periodic_thread

	active = True

	update_opacity(100)
	previous_image = None
	obs.timer_remove(update_ui)
	obs.timer_add(update_ui, int(refresh_interval * 1000))
	periodic_thread = continuous_threading.PeriodicThread(refresh_interval, update_backend)
	periodic_thread.start()
	

def deactivate_timer():
	global periodic_thread
	global active

	if periodic_thread != None:
		periodic_thread.stop()
		periodic_thread = None

	obs.timer_remove(update_ui)

	set_filter_value(screen_sourcename, "Color Correction", "opacity", 100)
	set_filter_value(camera_sourcename, "Blur", "Filter.Blur.Size", 0)

	active = False

def get_current_scene_name():
	scene = obs.obs_frontend_get_current_scene()
	scene_name = obs.obs_source_get_name(scene)
	obs.obs_source_release(scene)
	return scene_name

def on_event(event):
	if (event == obs.OBS_FRONTEND_EVENT_STREAMING_STARTED or event == obs.OBS_FRONTEND_EVENT_RECORDING_STARTED) and get_current_scene_name() == slide_scene and not active:
		activate_timer()
	elif event == obs.OBS_FRONTEND_EVENT_SCENE_CHANGED:
		if get_current_scene_name() == slide_scene:
			if obs.obs_frontend_streaming_active() or obs.obs_frontend_recording_active():
				if not active:
					activate_timer()
			else:
				set_filter_value(screen_sourcename, "Color Correction", "opacity", 100)
				set_filter_value(camera_sourcename, "Blur", "Filter.Blur.Size", 0)
		elif active:
			deactivate_timer()
	elif (event == obs.OBS_FRONTEND_EVENT_STREAMING_STOPPED or event == obs.OBS_FRONTEND_EVENT_RECORDING_STOPPED) and not (obs.obs_frontend_streaming_active() or obs.obs_frontend_recording_active()) and active:
		deactivate_timer()

def script_description():
	return "An OBS script to toggle the visiblility of a source for the purposes of a slide presentation.\nv" + version

def script_properties():
	props = obs.obs_properties_create()

	p = obs.obs_properties_add_list(props, "slide_scene", "Slide Scene", obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_STRING)
	obs.obs_property_list_add_string(p, "--Disabled--", "")
	scene_names = obs.obs_frontend_get_scene_names()
	if scene_names != None:
		for scene_name in scene_names:
			obs.obs_property_list_add_string(p, scene_name, scene_name)

	p = obs.obs_properties_add_list(props, "monitor", "Monitor", obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_INT)
	for i, monitor in enumerate(monitors):
		obs.obs_property_list_add_int(p, str(monitor.szDevice), i)

	p = obs.obs_properties_add_list(props, "screen_sourcename", "Screen Source", obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_STRING)
	obs.obs_property_list_add_string(p, "--Disabled--", "")
	sources = obs.obs_enum_sources()
	if sources != None:
		for source in sources:
			name = obs.obs_source_get_name(source)
			obs.obs_property_list_add_string(p, name, name)
	obs.source_list_release(sources)

	obs.obs_properties_add_int_slider(props, "slide_visible_duration", "Slide Visible Duration", 5, 120, 5)	

	obs.obs_properties_add_float_slider(props, "fadeout_duration", "Fade Out Duration", 0.05, 1.25, 0.05)

	obs.obs_properties_add_float_slider(props, "refresh_interval", "Refresh Interval", 0.1, 5, 0.1)

	obs.obs_properties_add_int_slider(props, "camera_blur", "Camera Blur", 1, 128, 1)

	p = obs.obs_properties_add_list(props, "camera_sourcename", "Camera Source", obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_STRING)
	obs.obs_property_list_add_string(p, "--Disabled--", "")
	sources = obs.obs_enum_sources()
	if sources != None:
		for source in sources:
			name = obs.obs_source_get_name(source)
			obs.obs_property_list_add_string(p, name, name)
	obs.source_list_release(sources)

	return props

def script_defaults(settings):
	obs.obs_data_set_default_int(settings, "slide_visible_duration", slide_visible_duration)
	obs.obs_data_set_default_double(settings, "fadeout_duration", fadeout_duration)
	obs.obs_data_set_default_double(settings, "refresh_interval", refresh_interval)
	obs.obs_data_set_default_int(settings, "camera_blur", camera_blur)

def script_update(settings):	
	global g
	g.settings = settings

	global slide_scene
	slide_scene = obs.obs_data_get_string(settings, "slide_scene")

	global monitors
	monitors = []
	for hMonitor, hdcMonitor, pyRect in win32api.EnumDisplayMonitors():
		monitors.append(Monitor(hMonitor, hdcMonitor, pyRect, win32api.GetMonitorInfo(hMonitor)["Device"])) 

	global monitor
	monitor = obs.obs_data_get_int(settings, "monitor")

	global screen_sourcename
	screen_sourcename = obs.obs_data_get_string(settings, "screen_sourcename")

	global slide_visible_duration
	slide_visible_duration = obs.obs_data_get_int(settings, "slide_visible_duration")

	global fadeout_duration
	fadeout_duration = obs.obs_data_get_double(settings, "fadeout_duration")

	global refresh_interval
	refresh_interval = obs.obs_data_get_double(settings, "refresh_interval")

	global camera_sourcename
	camera_sourcename = obs.obs_data_get_string(settings, "camera_sourcename")

def script_load(settings):
	global hotkey
	hotkey = obs.obs_hotkey_register_frontend("SlideDisplay.hotkey.show", "Show Slide-Display (hold)", hotkey_callback)
	hotkey_save_array = obs.obs_data_get_array(settings, "SlideDisplay.hotkey.show")
	obs.obs_hotkey_load(hotkey, hotkey_save_array)
	obs.obs_data_array_release(hotkey_save_array)

	obs.obs_frontend_add_event_callback(on_event)

def script_save(settings):
	hotkey_save_array = obs.obs_hotkey_save(hotkey)
	obs.obs_data_set_array(settings, "SlideDisplay.hotkey.show", hotkey_save_array)
	obs.obs_data_array_release(hotkey_save_array)

def script_unload():
	deactivate_timer()
	del g.settings

class Monitor(object):
	def __init__(self, hMonitor, hdcMonitor, pyRect, szDevice):
		self.hMonitor = hMonitor
		self.hdcMonitor = hdcMonitor
		self.pyRect = pyRect
		self.szDevice = szDevice

	def __eq__(self, other):
		return int(self.hMonitor) == int(other.hMonitor)

	def __hash__(self):
		return hash(self.hMonitor)