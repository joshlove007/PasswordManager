import os
import ctypes
import atexit

__all__ = ['HIDException', 'DeviceInfo', 'Device', 'enumerate']

library_paths = (
    'libhidapi-hidraw.so',
    'libhidapi-hidraw.so.0',
    'libhidapi-libusb.so',
    'libhidapi-libusb.so.0',
    'libhidapi-iohidmanager.so',
    'libhidapi-iohidmanager.so.0',
    'libhidapi.dylib',
    'hidapi.dll',
    'libhidapi-0.dll'
)

for lib in library_paths:
    try:
        hidapi = ctypes.cdll.LoadLibrary(lib)
        break
    except OSError:
        pass
else:
    error = "Unable to load any of the following libraries:{}"\
        .format(' '.join(library_paths))
    raise ImportError(error)


hidapi.hid_init()
atexit.register(hidapi.hid_exit)


class HIDException(Exception):
    pass


class DeviceInfo(ctypes.Structure):
    def as_dict(self):
        ret = {}
        for name, type in self._fields_:
            if name == 'next':
                continue
            ret[name] = getattr(self, name, None)

        return ret

DeviceInfo._fields_ = [
    ('path', ctypes.c_char_p),
    ('vendor_id', ctypes.c_ushort),
    ('product_id', ctypes.c_ushort),
    ('serial_number', ctypes.c_wchar_p),
    ('release_number', ctypes.c_ushort),
    ('manufacturer_string', ctypes.c_wchar_p),
    ('product_string', ctypes.c_wchar_p),
    ('usage_page', ctypes.c_ushort),
    ('usage', ctypes.c_ushort),
    ('interface_number', ctypes.c_int),
    ('next', ctypes.POINTER(DeviceInfo)),
]

hidapi.hid_init.argtypes = []
hidapi.hid_init.restype = ctypes.c_int
hidapi.hid_exit.argtypes = []
hidapi.hid_exit.restype = ctypes.c_int
hidapi.hid_enumerate.argtypes = [ctypes.c_ushort, ctypes.c_ushort]
hidapi.hid_enumerate.restype = ctypes.POINTER(DeviceInfo)
hidapi.hid_free_enumeration.argtypes = [ctypes.POINTER(DeviceInfo)]
hidapi.hid_free_enumeration.restype = None
hidapi.hid_open.argtypes = [ctypes.c_ushort, ctypes.c_ushort, ctypes.c_wchar_p]
hidapi.hid_open.restype = ctypes.c_void_p
hidapi.hid_open_path.argtypes = [ctypes.c_char_p]
hidapi.hid_open_path.restype = ctypes.c_void_p
hidapi.hid_write.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_size_t]
hidapi.hid_write.restype = ctypes.c_int
hidapi.hid_read_timeout.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_size_t, ctypes.c_int]
hidapi.hid_read_timeout.restype = ctypes.c_int
hidapi.hid_read.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_size_t]
hidapi.hid_read.restype = ctypes.c_int
hidapi.hid_set_nonblocking.argtypes = [ctypes.c_void_p, ctypes.c_int]
hidapi.hid_set_nonblocking.restype = ctypes.c_int
hidapi.hid_send_feature_report.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_int]
hidapi.hid_send_feature_report.restype = ctypes.c_int
hidapi.hid_get_feature_report.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_size_t]
hidapi.hid_get_feature_report.restype = ctypes.c_int
hidapi.hid_close.argtypes = [ctypes.c_void_p]
hidapi.hid_close.restype = None
hidapi.hid_get_manufacturer_string.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_size_t]
hidapi.hid_get_manufacturer_string.restype = ctypes.c_int
hidapi.hid_get_product_string.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_size_t]
hidapi.hid_get_product_string.restype = ctypes.c_int
hidapi.hid_get_serial_number_string.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_size_t]
hidapi.hid_get_serial_number_string.restype = ctypes.c_int
hidapi.hid_get_indexed_string.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_wchar_p, ctypes.c_size_t]
hidapi.hid_get_indexed_string.restype = ctypes.c_int
hidapi.hid_error.argtypes = [ctypes.c_void_p]
hidapi.hid_error.restype = ctypes.c_wchar_p


def enumerate(vid=0, pid=0):
    ret = []
    info = hidapi.hid_enumerate(vid, pid)
    c = info

    while c:
        ret.append(c.contents.as_dict())
        c = c.contents.next

    hidapi.hid_free_enumeration(info)

    return ret


class Device(object):
    def __init__(self, vid=None, pid=None, serial=None, path=None):
        if path:
            self.__dev = hidapi.hid_open_path(path)
        elif serial:
            serial = ctypes.create_unicode_buffer(serial)
            self.__dev = hidapi.hid_open(vid, pid, serial)
        elif vid and pid:
            self.__dev = hidapi.hid_open(vid, pid, None)
        else:
            raise ValueError('specify vid/pid or path')

        if not self.__dev:
            raise HIDException('unable to open device')

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, exc_traceback):
        self.close()

    def __hidcall(self, function, *args, **kwargs):
        if not self.__dev:
            raise HIDException('device closed')

        ret = function(*args, **kwargs)

        if ret == -1:
            err = hidapi.hid_error(self.__dev)
            raise HIDException(err)
        return ret

    def __readstring(self, function, max_length=255):
        buf = ctypes.create_unicode_buffer(max_length)
        self.__hidcall(function, self.__dev, buf, max_length)
        return buf.value

    def write(self, data):
        return self.__hidcall(hidapi.hid_write, self.__dev, data, len(data))

    def read(self, size, timeout=None):
        data = ctypes.create_string_buffer(size)

        if timeout is None:
            size = self.__hidcall(hidapi.hid_read, self.__dev, data, size)
        else:
            size = self.__hidcall(
                hidapi.hid_read_timeout, self.__dev, data, size, timeout)

        return data.raw[:size]

    def send_feature_report(self, data):
        return self.__hidcall(hidapi.hid_send_feature_report,
                              self.__dev, data, len(data))

    def get_feature_report(self, report_id, size):
        data = ctypes.create_string_buffer(size)

        # Pass the id of the report to be read.
        data[0] = bytearray((report_id,))

        size = self.__hidcall(
            hidapi.hid_get_feature_report, self.__dev, data, size)
        return data.raw[:size]

    def close(self):
        if self.__dev:
            hidapi.hid_close(self.__dev)
            self.__dev = None

    def open_path(self, path):
        self.__dev = hidapi.hid_open_path(path)
        
    def set_nonblocking(self, value):
        self.__hidcall(hidapi.hid_set_nonblocking, self.__dev, value)

    @property
    def nonblocking(self):
        return getattr(self, '_nonblocking', 0)

    @nonblocking.setter
    def nonblocking(self, value):
        self.__hidcall(hidapi.hid_set_nonblocking, self.__dev, value)
        setattr(self, '_nonblocking', value)

    @property
    def manufacturer(self):
        return self.__readstring(hidapi.hid_get_manufacturer_string)

    @property
    def product(self):
        return self.__readstring(hidapi.hid_get_product_string)

    @property
    def serial(self):
        return self.__readstring(hidapi.hid_get_serial_number_string)

    def get_indexed_string(self, index, max_length=255):
        buf = ctypes.create_unicode_buffer(max_length)
        self.__hidcall(hidapi.hid_get_indexed_string,
                       self.__dev, index, buf, max_length)
        return buf.value


class device(object):    
    def open_path(self, path):
        self.__dev = hidapi.hid_open_path(path)
    
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, exc_traceback):
        self.close()

    def __hidcall(self, function, *args, **kwargs):
        if not self.__dev:
            raise HIDException('device closed')

        ret = function(*args, **kwargs)

        if ret == -1:
            err = hidapi.hid_error(self.__dev)
            raise HIDException(err)
        return ret

    def __readstring(self, function, max_length=255):
        buf = ctypes.create_unicode_buffer(max_length)
        self.__hidcall(function, self.__dev, buf, max_length)
        return buf.value

    def write(self, buff):
        data = bytes(buff)
        return self.__hidcall(hidapi.hid_write, self.__dev, data, len(data))

    def read(self, max_length, timeout_ms=None):
        data = ctypes.create_string_buffer(max_length)

        if timeout_ms is None:
            max_length = self.__hidcall(hidapi.hid_read, self.__dev, data, max_length)
        else:
            max_length = self.__hidcall(
                hidapi.hid_read_timeout, self.__dev, data, max_length, timeout_ms)

        return data.raw[:max_length]

    def send_feature_report(self, data):
        return self.__hidcall(hidapi.hid_send_feature_report,
                              self.__dev, data, len(data))

    def get_feature_report(self, report_id, size):
        data = ctypes.create_string_buffer(size)

        # Pass the id of the report to be read.
        data[0] = bytearray((report_id,))

        size = self.__hidcall(
            hidapi.hid_get_feature_report, self.__dev, data, size)
        return data.raw[:size]

    def close(self):
        if self.__dev:
            hidapi.hid_close(self.__dev)
            self.__dev = None
        
    def set_nonblocking(self, value):
        self.__hidcall(hidapi.hid_set_nonblocking, self.__dev, value)
        setattr(self, '_nonblocking', value)

    @property
    def nonblocking(self):
        return getattr(self, '_nonblocking', 0)

    @nonblocking.setter
    def nonblocking(self, value):
        self.__hidcall(hidapi.hid_set_nonblocking, self.__dev, value)
        setattr(self, '_nonblocking', value)

    @property
    def manufacturer(self):
        return self.__readstring(hidapi.hid_get_manufacturer_string)

    @property
    def product(self):
        return self.__readstring(hidapi.hid_get_product_string)

    @property
    def serial(self):
        return self.__readstring(hidapi.hid_get_serial_number_string)

    def get_indexed_string(self, index, max_length=255):
        buf = ctypes.create_unicode_buffer(max_length)
        self.__hidcall(hidapi.hid_get_indexed_string,
                       self.__dev, index, buf, max_length)
        return buf.value

'''
close(...)
    Close connection.
    
    This should always be called after opening a connection.

error(...)
    Get error from device.
    
    :return:
    :rtype: str
    :raises ValueError: If connection is not opened.
    :raises IOError:

get_feature_report(...)
    Receive feature report.
    
    :param report_num:
    :type report_num: int
    :param max_length:
    :type max_length: int
    :return: Incoming feature report
    :rtype: List[int]
    :raises ValueError: If connection is not opened.
    :raises IOError:

get_indexed_string(...)
    Return indexed string.
    
    :return:
    :rtype: str
    :raises ValueError: If connection is not opened.
    :raises IOError:

get_input_report(...)
    Get input report
    
    :param report_num:
    :type report_num: int
    :param max_length:
    :type max_length: int
    :return:
    :rtype: List[int]
    :raises ValueError: If connection is not opened.
    :raises IOError:

get_manufacturer_string(...)
    Return manufacturer string (e.g. vendor name).
    
    :return:
    :rtype: str
    :raises ValueError: If connection is not opened.
    :raises IOError:

get_product_string(...)
    Return product string (e.g. device description).
    
    :return:
    :rtype: str
    :raises ValueError: If connection is not opened.
    :raises IOError:

get_serial_number_string(...)
    Return serial number.
    
    :return:
    :rtype: str
    :raises ValueError: If connection is not opened.
    :raises IOError:

open(...)
    Open the connection.
    
    :param vendor_id: Vendor id to connect to, default = 0
    :type vendor_id: int, optional
    :param product_id: Product id to connect to, default = 0
    :type product_id: int, optional
    :param serial_number:
    :type serial_number: unicode, optional
    :raises IOError:

open_path(...)
    Open connection by path.
    
    :param path: Path to device
    :type path: bytes
    :raises IOError:

read(...)
    Return a list of integers (0-255) from the device up to max_length bytes.
    
    :param max_length: Maximum number of bytes to read
    :type max_length: int
    :param timeout_ms: Number of milliseconds until timeout (default: no timeout)
    :type timeout_ms: int, optional
    :return: Read bytes
    :rtype: List[int]

send_feature_report(...)
    Accept a list of integers (0-255) and send them to the device.
    
    :param buff: Data to send (must be convertible into bytes)
    :type buff: any
    :return: Send result
    :rtype: int

set_nonblocking(...)
    Set the nonblocking flag.
    
    :param v: Flag value (1 or 0, True or False)
    :type v: int, bool
    :return: Flag result
    :rtype: int

write(...)
    Accept a list of integers (0-255) and send them to the device.
    
    :param buff: Data to write (must be convertible to `bytes`)
    :type buff: Any
    :return: Write result
    :rtype: int

----------------------------------------------------------------------
Static methods defined here:

__new__(*args, **kwargs) from builtins.type
    Create and return a new object.  See help(type) for accurate signature.
'''