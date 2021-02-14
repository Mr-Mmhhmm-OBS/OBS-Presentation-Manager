import obspython as obs
import win32api
from PIL import ImageGrab
import time
import continuous_threading

version = "1.5"

monitors = []
monitor = None

slide_scenes = []
active = False

screen_source = None
camera_source = None
camera_blur = 1
screen_visible = True

previous_image = []
timestamp = 0

slide_visible_duration = 10

refresh_interval = 0.1
periodic_thread = None

def update_opacity(source_name, value):
	global screen_visible
	
	screen_visible = (value == 100)

	update_filter(screen_source, "Color Correction", "opacity", value)
	update_filter(camera_source, "Blur", "Filter.Blur.Size", camera_blur if screen_visible else 1)

def update_filter(source_name, filter_name, filter_field_name, value):
	source = obs.obs_get_source_by_name(source_name)
	if source is not None:
		filter = obs.obs_source_get_filter_by_name(source, filter_name)
		if filter is not None:
			# Get the settings data object for the filter
			filter_settings = obs.obs_source_get_settings(filter)

			# Update the hue_shift property and update the filter with the new settings
			obs.obs_data_set_double(filter_settings, filter_field_name, value)
			obs.obs_source_update(filter, filter_settings)

			# Release the resources
			obs.obs_data_release(filter_settings)
			obs.obs_source_release(filter)
		obs.obs_source_release(source)

def update():
	global previous_image
	global timestamp

	im = ImageGrab.grab(monitors[monitor].pyRect, False, True, None)

	if im != previous_image:
		previous_image = im
		if im.convert("L").getextrema() == (0,0):
			update_opacity(screen_source, 0)
		elif not screen_visible:
			update_opacity(screen_source, 100)
			timestamp = time.time()
	elif time.time() >= timestamp + slide_visible_duration and screen_visible:
			update_opacity(screen_source, 0)

def activate_timer():
	global active
	global previous_image
	global timestamp
	global periodic_thread

	timestamp = time.time()
	update_opacity(screen_source, 100)
	active = True
	previous_image = None
	periodic_thread = continuous_threading.PeriodicThread(0.1, update)
	periodic_thread.start()
	

def deactivate_timer():
	global active

	if periodic_thread != None:
		periodic_thread.join()

	update_opacity(screen_source, 100)
	active = False

def get_current_scene_name():
	scene = obs.obs_frontend_get_current_scene()
	scene_name = obs.obs_source_get_name(scene)
	obs.obs_source_release(scene);
	return scene_name

def on_event(event):
	if (event == obs.OBS_FRONTEND_EVENT_STREAMING_STARTED or event == obs.OBS_FRONTEND_EVENT_RECORDING_STARTED) and get_current_scene_name() in slide_scenes and not active:
		if get_current_scene_name() in slide_scenes:
			if not active:
				activate_timer()
	elif event == obs.OBS_FRONTEND_EVENT_SCENE_CHANGED and (obs.obs_frontend_streaming_active() or obs.obs_frontend_recording_active()):
		if get_current_scene_name() in slide_scenes:
			if not active:
				activate_timer()
		else:
			deactivate_timer()
	elif (event == obs.OBS_FRONTEND_EVENT_STREAMING_STOPPED or event == obs.OBS_FRONTEND_EVENT_RECORDING_STOPPED) and active:
		deactivate_timer()

def script_description():
	return "An OBS script to toggle the visiblility of a source for the purposes of a slide presentation.\nv" + version

def script_properties():
	props = obs.obs_properties_create()

	p = obs.obs_properties_add_list(props, "monitor", "Monitor", obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_INT)
	for i, monitor in enumerate(monitors):
		obs.obs_property_list_add_int(p, str(monitor.szDevice), i)

	group = obs.obs_properties_create()
	scene_names = obs.obs_frontend_get_scene_names()
	if scene_names != None:
		for scene_name in scene_names:
			obs.obs_properties_add_bool(group, "slide_scene_" + str(scene_name), scene_name)
	obs.obs_properties_add_group(props, "slide_scenes", "Slide Scenes", obs.OBS_GROUP_NORMAL, group)

	screen_source_list = obs.obs_properties_add_list(props, "screen_source", "Screen Source", obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_STRING)
	obs.obs_property_list_add_string(screen_source_list, "--Disabled--", "")
	sources = obs.obs_enum_sources()
	if sources != None:
		for source in sources:
			if obs.obs_source_get_unversioned_id(source) == "monitor_capture":
				name = obs.obs_source_get_name(source)
				obs.obs_property_list_add_string(screen_source_list, name, name)
	obs.source_list_release(sources)

	screen_source_list = obs.obs_properties_add_list(props, "camera_source", "Camera Source", obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_STRING)
	obs.obs_property_list_add_string(screen_source_list, "--Disabled--", "")
	sources = obs.obs_enum_sources()
	if sources != None:
		for source in sources:
			if obs.obs_source_get_unversioned_id(source) == "dshow_input":
				name = obs.obs_source_get_name(source)
				obs.obs_property_list_add_string(screen_source_list, name, name)
	obs.source_list_release(sources)

	obs.obs_properties_add_int_slider(props, "camera_blur", "Camera Blur", 1, 127, camera_blur)

	obs.obs_properties_add_int_slider(props, "slide_visible_duration", "Slide Visible Duration", 5, 120, 5)

	obs.obs_properties_add_float_slider(props, "refresh_interval", "Refresh Interval", 0.1, 5, 0.1)

	return props

def script_defaults(settings):
	obs.obs_data_set_default_int(settings, "slide_visible_duration", slide_visible_duration)
	obs.obs_data_set_default_double(settings, "refresh_interval", refresh_interval)
	obs.obs_data_set_default_int(settings, "camera_blur", camera_blur)

def script_update(settings):
	global monitors
	monitors = []
	for hMonitor, hdcMonitor, pyRect in win32api.EnumDisplayMonitors():
		monitors.append(Monitor(hMonitor, hdcMonitor, pyRect, win32api.GetMonitorInfo(hMonitor)["Device"])) 

	global monitor
	monitor = obs.obs_data_get_int(settings, "monitor")


	scene_names = obs.obs_frontend_get_scene_names()
	if scene_names != None and len(scene_names) > 0:
		# Update scene_name list
		array = obs.obs_data_array_create()
		for i, scene_name in enumerate(scene_names):
			data_item = obs.obs_data_create()
			obs.obs_data_set_string(data_item, "scene_name", scene_name)
			obs.obs_data_array_insert(array, i, data_item)
		obs.obs_data_set_array(settings, "scene_names", array)

	global slide_scenes
	slide_scenes = []
	scene_name_array = obs.obs_data_get_array(settings, "scene_names")
	if scene_name_array != None:
		for i in range(obs.obs_data_array_count(scene_name_array)):
			data_item = obs.obs_data_array_item(scene_name_array, i)
			scene_name = obs.obs_data_get_string(data_item, "scene_name")
			checked = obs.obs_data_get_bool(settings, "slide_scene_" + str(scene_name))
			if checked:
				slide_scenes.append(scene_name)
		obs.obs_data_array_release(scene_name_array)

	global screen_source
	screen_source = obs.obs_data_get_string(settings, "screen_source")

	global camera_source
	camera_source = obs.obs_data_get_string(settings, "camera_source")

	global camera_blur
	camera_blur = obs.obs_data_get_int(settings, "camera_blur")

	global slide_visible_duration
	slide_visible_duration = obs.obs_data_get_int(settings, "slide_visible_duration")

	global refresh_interval
	refresh_interval = obs.obs_data_get_double(settings, "refresh_interval")
	
	deactivate_timer()

	obs.obs_frontend_add_event_callback(on_event)

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

	def width(self):
		return self.pyRect[2] - self.pyRect[0]
	
	def height(self):
		return self.pyRect[3] - self.pyRect[1]