import obspython as obs
import sys
import win32api
from PIL import ImageGrab
import time
import continuous_threading

version = "1.8"

g = lambda: ...
g.settings = None

monitors = []
monitor = None

slide_scenes = []
active = False

filters = []

screen_visible = True

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
	global timestamp
	global status
	status = DEFAULT_STATUS

	value = int(value)

	screen_visible = (value == 100)
	if screen_visible:
		obs.timer_remove(fadeout_callback)
	timestamp = time.time() if screen_visible else 0

	for filter in filters:
		if filter.source_name != "":
			set_filter_value(filter.source_name, filter.filter_name, filter.filterfield_name, int((filter.on_value - filter.off_value) * (value / 100) + filter.off_value))

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

	fadeout_timestamp = time.time()
	obs.timer_add(fadeout_callback, 50)

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

	update_opacity(100)
	active = True
	previous_image = None
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
	active = False

def get_current_scene_name():
	scene = obs.obs_frontend_get_current_scene()
	scene_name = obs.obs_source_get_name(scene)
	obs.obs_source_release(scene);
	return scene_name

def SetDefaultFilterValues():
	for filter in filters:
		if filter.source_name != "":
			set_filter_value(filter.source_name, filter.filter_name, filter.filterfield_name, filter.default_value)

def on_event(event):
	if (event == obs.OBS_FRONTEND_EVENT_STREAMING_STARTED or event == obs.OBS_FRONTEND_EVENT_RECORDING_STARTED) and get_current_scene_name() in slide_scenes and not active:
		if get_current_scene_name() in slide_scenes:
			if not active:
				activate_timer()
	elif event == obs.OBS_FRONTEND_EVENT_SCENE_CHANGED:
		if get_current_scene_name() in slide_scenes:
			if obs.obs_frontend_streaming_active() or obs.obs_frontend_recording_active():
				if not active:
					activate_timer()
			else:
				SetDefaultFilterValues()
		else:
			if active:
				deactivate_timer()
			SetDefaultFilterValues()
	elif (event == obs.OBS_FRONTEND_EVENT_STREAMING_STOPPED or event == obs.OBS_FRONTEND_EVENT_RECORDING_STOPPED) and active:
		deactivate_timer()

def script_description():
	return "An OBS script to toggle the visiblility of a source for the purposes of a slide presentation.\nv" + version

def update_filter_data(i, source_name, filter_name, fitlerfield_name, default_value, off_value, on_value):
	obs.obs_data_set_string(g.settings, "filter"+str(i)+"_sourcename", source_name)
	obs.obs_data_set_string(g.settings, "filter"+str(i)+"_filtername", filter_name)
	obs.obs_data_set_string(g.settings, "filter"+str(i)+"_filterfieldname", fitlerfield_name)
	obs.obs_data_set_int(g.settings, "filter"+str(i)+"_defaultvalue", default_value)
	obs.obs_data_set_int(g.settings, "filter"+str(i)+"_offvalue", off_value)
	obs.obs_data_set_int(g.settings, "filter"+str(i)+"_onvalue", on_value)

def erase_filter_data(i):
	obs.obs_data_erase(g.settings, "filter"+str(i)+"_sourcename")
	obs.obs_data_erase(g.settings, "filter"+str(i)+"_filtername")
	obs.obs_data_erase(g.settings, "filter"+str(i)+"_filterfieldname")
	obs.obs_data_erase(g.settings, "filter"+str(i)+"_defaultvalue")
	obs.obs_data_erase(g.settings, "filter"+str(i)+"_offvalue")
	obs.obs_data_erase(g.settings, "filter"+str(i)+"_onvalue")


def addfilter_callback(props, property):
	i = len(filters)
	update_filter_data(i, "", "", "", 0, 0 ,0)
	filters.append(OBS_Filter("", "", "", 0, 0, 0))
	obs.obs_data_set_int(g.settings, "filters", len(filters))
	
	group = obs.obs_properties_get(props, "filter"+str(i)+"_group")
	if group is None:
		add_filter_group(obs.obs_property_group_content(obs.obs_properties_get(props, "filter_groups")), i)
	else:
		obs.obs_property_set_visible(group, True)
	return True

def removefilter_callback(props, property):
	i = int(obs.obs_property_name(property).split('_')[0].split('filter')[1])

	if (i < len(filters)):
		filters.pop(i)
		erase_filter_data(len(filters))
		obs.obs_property_set_visible(obs.obs_properties_get(props, "filter"+str(len(filters))+"_group"), False)
		obs.obs_data_set_int(g.settings, "filters", len(filters))
		for i2 in range(i, len(filters)):
			update_filter_data(i2, filters[i2].source_name, filters[i2].filter_name, filters[i2].filterfield_name, filters[i2].default_value, filters[i2].off_value, filters[i2].on_value)
	return True

def filter_sourcename_modified_callback(props, property, settings):
	index = int(obs.obs_property_name(property).split('_')[0].split('filter')[1])
	filters[index].source_name = obs.obs_data_get_string(settings, "filter"+str(index)+"_sourcename")

def filter_filtername_modified_callback(props, property, settings):
	index = int(obs.obs_property_name(property).split('_')[0].split('filter')[1])
	filters[index].filter_name = obs.obs_data_get_string(settings, "filter"+str(index)+"_filtername")

def filter_filterfieldname_modified_callback(props, property, settings):
	index = int(obs.obs_property_name(property).split('_')[0].split('filter')[1])
	filters[index].filterfield_name = obs.obs_data_get_string(settings, "filter"+str(index)+"_filterfieldname")

def filter_defaultvalue_modified_callback(props, property, settings):
	index = int(obs.obs_property_name(property).split('_')[0].split('filter')[1])
	filters[index].default_value = obs.obs_data_get_int(settings, "filter"+str(index)+"_defaltvalue")

def filter_offvalue_modified_callback(props, property, settings):
	index = int(obs.obs_property_name(property).split('_')[0].split('filter')[1])
	filters[index].off_value = obs.obs_data_get_int(settings, "filter"+str(index)+"_offvalue")

def filter_onvalue_modified_callback(props, property, settings):
	index = int(obs.obs_property_name(property).split('_')[0].split('filter')[1])
	filters[index].on_value = obs.obs_data_get_int(settings, "filter"+str(index)+"_onvalue")

def add_filter_group(filter_groups, i):
	filter_group = obs.obs_properties_create()

	prop = obs.obs_properties_add_list(filter_group, "filter"+str(i)+"_sourcename", "Source", obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_STRING)
	obs.obs_property_set_modified_callback(prop, filter_sourcename_modified_callback)
	obs.obs_property_list_add_string(prop, "--Disabled--", "")
	sources = obs.obs_enum_sources()
	if sources != None:
		for source in sources:
			name = obs.obs_source_get_name(source)
			obs.obs_property_list_add_string(prop, name, name)
	obs.source_list_release(sources)

	prop = obs.obs_properties_add_text(filter_group, "filter"+str(i)+"_filtername", "Filter Name", obs.OBS_TEXT_DEFAULT)
	obs.obs_property_set_modified_callback(prop, filter_filtername_modified_callback)
	prop = obs.obs_properties_add_text(filter_group, "filter"+str(i)+"_filterfieldname", "Filter Field Name", obs.OBS_TEXT_DEFAULT)
	obs.obs_property_set_modified_callback(prop, filter_filterfieldname_modified_callback)
	prop = obs.obs_properties_add_int(filter_group, "filter"+str(i)+"_defaultvalue", "Default Value", -256, 256, 1)
	obs.obs_property_set_modified_callback(prop, filter_defaultvalue_modified_callback)
	prop = obs.obs_properties_add_int(filter_group, "filter"+str(i)+"_offvalue", "Off Value", -256, 256, 1)
	obs.obs_property_set_modified_callback(prop, filter_offvalue_modified_callback)
	prop = obs.obs_properties_add_int(filter_group, "filter"+str(i)+"_onvalue", "On Value", -256, 256, 1)
	obs.obs_property_set_modified_callback(prop, filter_onvalue_modified_callback)
	
	obs.obs_properties_add_button(filter_group, "filter"+str(i)+"_removefilter", "Remove Filter", removefilter_callback)

	obs.obs_properties_add_group(filter_groups, "filter"+str(i)+"_group", "Filter #"+str(i+1), obs.OBS_GROUP_NORMAL, filter_group)

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

	obs.obs_properties_add_int_slider(props, "slide_visible_duration", "Slide Visible Duration", 5, 120, 5)	

	obs.obs_properties_add_float_slider(props, "fadeout_duration", "Fade Out Duration", 0.05, 1.25, 0.05)

	obs.obs_properties_add_float_slider(props, "refresh_interval", "Refresh Interval", 0.1, 5, 0.1)

	filter_groups = obs.obs_properties_create()
	for i, filter in enumerate(filters):
		add_filter_group(filter_groups, i)
	obs.obs_properties_add_group(props, "filter_groups", "Filters", obs.OBS_GROUP_NORMAL, filter_groups)
	obs.obs_properties_add_button(props, "add_filter", "Add Filter", addfilter_callback)

	return props

def script_defaults(settings):
	obs.obs_data_set_default_int(settings, "slide_visible_duration", slide_visible_duration)	
	obs.obs_data_set_default_double(settings, "fadeout_duration", fadeout_duration)
	obs.obs_data_set_default_double(settings, "refresh_interval", refresh_interval)

def script_update(settings):	
	global g
	g.settings = settings

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

	global slide_visible_duration
	slide_visible_duration = obs.obs_data_get_int(settings, "slide_visible_duration")

	global fadeout_duration
	fadeout_duration = obs.obs_data_get_double(settings, "fadeout_duration")

	global refresh_interval
	refresh_interval = obs.obs_data_get_double(settings, "refresh_interval")	

	global filters
	filters.clear()
	for i in range(obs.obs_data_get_int(settings, "filters")):
		filters.append(OBS_Filter(obs.obs_data_get_string(settings, "filter"+str(i)+"_sourcename"), 
			obs.obs_data_get_string(settings, "filter"+str(i)+"_filtername"), 
			obs.obs_data_get_string(settings, "filter"+str(i)+"_filterfieldname"), 
			obs.obs_data_get_int(settings, "filter"+str(i)+"_defaultvalue"), 
			obs.obs_data_get_int(settings, "filter"+str(i)+"_offvalue"), 
			obs.obs_data_get_int(settings, "filter"+str(i)+"_onvalue"))
		)

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

	def width(self):
		return self.pyRect[2] - self.pyRect[0]
	
	def height(self):
		return self.pyRect[3] - self.pyRect[1]

class OBS_Filter(object):
	def __init__(self, source_name, filter_name, filterfield_name, default_value, off_value, on_value):
		self.source_name = source_name
		self.filter_name = filter_name
		self.filterfield_name = filterfield_name
		self.default_value = default_value
		self.off_value = off_value
		self.on_value = on_value

	def __str__(self):
		return self.source_name