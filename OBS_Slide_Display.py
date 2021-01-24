import obspython as obs
import win32api
from PIL import ImageGrab
import time

version = "1.2"

monitors = []
monitor = None

slide_scenes = []
active = False

screen_source = None
screen_visible = True

#resolution = 300
previous_image = []
timestamp = 0

slide_visible_duration = 10

refresh_interval = 500

def show_source(source_name, show):
	current_scene = obs.obs_frontend_get_current_scene()
	current_scene = obs.obs_scene_from_source(current_scene)
	sceneitems = obs.obs_scene_enum_items(current_scene)
	if sceneitems != None:
		for sceneitem in sceneitems:
			source = obs.obs_sceneitem_get_source(sceneitem)
			if obs.obs_source_get_name(source) == source_name:
				obs.obs_sceneitem_set_visible(sceneitem, show)
				break
			obs.obs_source_release(source)
		obs.sceneitem_list_release(sceneitems)

def update():
	global previous_image
	global timestamp
	global screen_visible

	im = ImageGrab.grab(monitors[monitor].pyRect, False, True, None)

	#height = monitors[monitor].height()
	#hstep = int(height / resolution)
	#width = monitors[monitor].width()
	#wstep = int(width / resolution)

	#pixels = []
	#for y in range(hstep, height, hstep):
	#	for x in range(wstep, width, wstep):
	#		pixels.append(im.getpixel((x,y)))

	if im != previous_image:
		previous_image = im
		show_source(screen_source, True)
		timestamp = time.time()
		screen_visible = True
	elif time.time() - timestamp > slide_visible_duration and screen_visible:
		print("too old")
		show_source(screen_source, False)
		screen_visible = False

def activate_timer():
	global active
	obs.timer_add(update, refresh_interval)
	active = True

def deactivate_timer():
	global active
	obs.timer_remove(update)
	active = False

def get_current_scene_name():
	scene = obs.obs_frontend_get_current_scene()
	scene_name = obs.obs_source_get_name(scene)
	obs.obs_source_release(scene);
	return scene_name

def on_event(event):
	if (event == obs.OBS_FRONTEND_EVENT_STREAMING_STARTED or event == obs.OBS_FRONTEND_EVENT_RECORDING_STARTED) and get_current_scene_name() in slide_scenes and not active:
		activate_timer()
	elif event == obs.OBS_FRONTEND_EVENT_SCENE_CHANGED and (obs.obs_frontend_streaming_active() or obs.obs_frontend_recording_active()):
		if get_current_scene_name() in slide_scenes: #TODO: correct scene
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

	obs.obs_properties_add_int_slider(props, "slide_visible_duration", "Slide Visible Duration", 5, 120, 5)

	obs.obs_properties_add_int_slider(props, "refresh_interval", "Refresh Interval", 100, 5000, 100)

	return props

def script_defaults(settings):
	obs.obs_data_set_default_int(settings, "slide_visible_duration", slide_visible_duration)
	obs.obs_data_set_default_int(settings, "refresh_interval", refresh_interval)

def script_update(settings):
	global monitors
	monitors = []
	for hMonitor, hdcMonitor, pyRect in win32api.EnumDisplayMonitors():
		monitors.append(Monitor(hMonitor, hdcMonitor, pyRect, win32api.GetMonitorInfo(hMonitor)["Device"])) 

	global monitor
	monitor = obs.obs_data_get_int(settings, "monitor")

	global slide_scenes
	for scene_name in obs.obs_frontend_get_scene_names():
		if obs.obs_data_get_bool(settings, "slide_scene_" + str(scene_name)):
			slide_scenes.append(scene_name)

	global screen_source
	screen_source = obs.obs_data_get_string(settings, "screen_source")

	global slide_visible_duration
	slide_visible_duration = obs.obs_data_get_int(settings, "slide_visible_duration")

	global refresh_interval
	refresh_interval = obs.obs_data_get_int(settings, "refresh_interval")

def script_load(settings):
	obs.timer_remove(update)
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