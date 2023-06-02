#!/usr/bin/env python
"""
oleobj.py

oleobj是一个Python脚本和模块，用于解析存储的OLE对象和文件
转换为各种MS Office文件格式（doc、xls、ppt、docx、xlsx、pptx等）

作者: Philippe Lagadec（菲利普·拉加德克） - http://www.decalage.info
许可证：BSD，请参阅源代码或文档

oleobj是python-oletools包的一部分：
http://www.decalage.info/python/oletools
"""

# === 许可证 =================================================================

# oleobj is copyright (c) 2015-2022 Philippe Lagadec (http://www.decalage.info)
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
#  * Redistributions of source code must retain the above copyright notice,
#    this list of conditions and the following disclaimer.
#  * Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
# ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
# LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
# CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
# SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
# INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
# ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
# POSSIBILITY OF SUCH DAMAGE.


# -- 导入文件 ------------------------------------------------------------------

from __future__ import print_function

import logging
import struct
import argparse
import os
import re
import sys
import io
from zipfile import is_zipfile
import random

import olefile

# 重要: 应该可以在任何目录中直接作为脚本运行oletools，
# 而无需使用pip或setup.py安装它们。
# 在这种情况下，相对导入不可用。
# 为了实现Python 2+3兼容性，我们需要使用绝对导入，
# 因此我们将oletools父文件夹添加到sys.path（绝对+规范化路径）：
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
# print('_thismodule_dir = %r' % _thismodule_dir)
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))
# print('_parent_dir = %r' % _thirdparty_dir)
if _parent_dir not in sys.path:
    sys.path.insert(0, _parent_dir)

from oletools.thirdparty import xglob
from oletools.ppt_record_parser import (is_ppt, PptFile,
                                        PptRecordExOleVbaActiveXAtom)
from oletools.ooxml import XmlParser
from oletools.common.io_encoding import ensure_stdout_handles_unicode

# -----------------------------------------------------------------------------
# 修改日志:
# 2015-12-05 v0.01 PL: - 第一个版本
# 2016-06          PL: - 添加了main和process_file（尚未工作）
# 2016-07-18 v0.48 SL: - 添加了 Python 3.5 支持
# 2016-07-19       PL: - 修复了 Python 2.6-7 支持
# 2016-11-17 v0.51 PL: - 修复了 OLE 原始对象提取
# 2016-11-18       PL: - 为setup.py入口点添加了main
# 2017-05-03       PL: - 修复了绝对导入 (issue #141)
# 2018-01-18 v0.52 CH: - 添加了对压缩xml的类型的支持 (docx, pptx,
#                        xlsx), and ppt
# 2018-03-27       PL: - 修复了 issue #274 在 read_length_prefixed_string 中
# 2018-09-11 v0.54 PL: - olefile 现在是一个依赖项
# 2018-10-30       SA: - 添加了对外部链接的检测 (PR #317)
# 2020-03-03 v0.56 PL: - 修复了错误#541，“Ole10Native”不区分大小写
# 2022-01-28 v0.60 PL: - 添加了对customUI标记的检测

__version__ = '0.60.1'

# -----------------------------------------------------------------------------
# TODO:
# + 设置日志记录（与其他oletools通用）


# -----------------------------------------------------------------------------
# 参考文档:

# 嵌入式OLE对象/文件的存储参考：
# [MS-OLEDS]: 对象链接和嵌入（OLE）数据结构
# https://msdn.microsoft.com/en-us/library/dd942265.aspx

# - office 分析器: https://github.com/unixfreak0037/officeparser
# TODO: ole转存


# === 日志 =================================================================

DEFAULT_LOG_LEVEL = "warning"
LOG_LEVELS = {'debug':    logging.DEBUG,
              'info':     logging.INFO,
              'warning':  logging.WARNING,
              'error':    logging.ERROR,
              'critical': logging.CRITICAL,
              'debug-olefile': logging.DEBUG}


class NullHandler(logging.Handler):
    """
    没有输出的日志处理程序，
    以避免在主应用程序未配置日志记录时打印消息。
    Python 2.7有logging.NullHandler，但这对于2.6来说是必要的:
    查看文档： https://docs.python.org/2.6/library/logging.html
    configuring-logging-for-a-library
    """
    def emit(self, record):
        pass


def get_logger(name, level=logging.CRITICAL+1):
    """
    为此模块创建一个合适的日志对象。
    目标不是更改根日志对象的设置，以避免
    其他模块的日志显示在屏幕上。
    如果存在具有相同名称的日志对象，请重用它。 (否则，它将具有重复的处理程序，
    并且消息将加倍。)
    默认情况下，该级别设置为CRITICAL+1，以避免任何日志记录。
    """
    # 首先，测试是否已经有一个具有相同名称的日志对象，
    # 否则它将生成重复的消息（由于重复的处理程序）：
    if name in logging.Logger.manager.loggerDict:
        # NOTE: another less intrusive but more "hackish" solution would be to
        # use getLogger then test if its effective level is not default.
        logger = logging.getLogger(name)
        # make sure level is OK:
        logger.setLevel(level)
        return logger
    # get a new logger:
    logger = logging.getLogger(name)
    # only add a NullHandler for this logger, it is up to the application
    # to configure its own logging:
    logger.addHandler(NullHandler())
    logger.setLevel(level)
    return logger


# a global logger object used for debugging:
log = get_logger('oleobj')     # pylint: disable=invalid-name


def enable_logging():
    """
    Enable logging for this module (disabled by default).
    This will set the module-specific logger level to NOTSET, which
    means the main application controls the actual logging level.
    """
    log.setLevel(logging.NOTSET)


# === CONSTANTS ===============================================================

# some str methods on Python 2.x return characters,
# while the equivalent bytes methods return integers on Python 3.x:
if sys.version_info[0] <= 2:
    # Python 2.x
    NULL_CHAR = '\x00'
else:
    # Python 3.x
    NULL_CHAR = 0     # pylint: disable=redefined-variable-type
    xrange = range    # pylint: disable=redefined-builtin, invalid-name

OOXML_RELATIONSHIP_TAG = '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'
# There are several customUI tags for different versions of Office:
TAG_CUSTOMUI_2007 = "{http://schemas.microsoft.com/office/2006/01/customui}customUI"
TAG_CUSTOMUI_2010 = "{http://schemas.microsoft.com/office/2009/07/customui}customUI"

# === GLOBAL VARIABLES ========================================================

# struct to parse an unsigned integer of 32 bits:
STRUCT_UINT32 = struct.Struct('<L')
assert STRUCT_UINT32.size == 4  # make sure it matches 4 bytes

# struct to parse an unsigned integer of 16 bits:
STRUCT_UINT16 = struct.Struct('<H')
assert STRUCT_UINT16.size == 2  # make sure it matches 2 bytes

# max length of a zero-terminated ansi string. Not sure what this really is
STR_MAX_LEN = 1024

# size of chunks to copy from ole stream to file
DUMP_CHUNK_SIZE = 4096

# return values from main; can be added
# (e.g.: did dump but had err parsing and dumping --> return 1+4+8 = 13)
RETURN_NO_DUMP = 0     # nothing found to dump/extract
RETURN_DID_DUMP = 1    # did dump/extract successfully
RETURN_ERR_ARGS = 2    # reserve for OptionParser.parse_args
RETURN_ERR_STREAM = 4  # error opening/parsing a stream
RETURN_ERR_DUMP = 8    # error dumping data from stream to file

# Not sure if they can all be "External", but just in case
BLACKLISTED_RELATIONSHIP_TYPES = [
    'attachedTemplate',
    'externalLink',
    'externalLinkPath',
    'externalReference',
    'frame',
    'hyperlink',
    'officeDocument',
    'oleObject',
    'package',
    'slideUpdateUrl',
    'slideMaster',
    'slide',
    'slideUpdateInfo',
    'subDocument',
    'worksheet'
]

# Save maximum length of a filename
MAX_FILENAME_LENGTH = 255

# Max attempts at generating a non-existent random file name
MAX_FILENAME_ATTEMPTS = 100

# === FUNCTIONS ===============================================================


def read_uint32(data, index):
    """
    Read an unsigned integer from the first 32 bits of data.

    :param data: bytes string or stream containing the data to be extracted.
    :param index: index to start reading from or None if data is stream.
    :return: tuple (value, index) containing the read value (int),
             and the index to continue reading next time.
    """
    if index is None:
        value = STRUCT_UINT32.unpack(data.read(4))[0]
    else:
        value = STRUCT_UINT32.unpack(data[index:index+4])[0]
        index += 4
    return (value, index)


def read_uint16(data, index):
    """
    Read an unsigned integer from the 16 bits of data following index.

    :param data: bytes string or stream containing the data to be extracted.
    :param index: index to start reading from or None if data is stream
    :return: tuple (value, index) containing the read value (int),
             and the index to continue reading next time.
    """
    if index is None:
        value = STRUCT_UINT16.unpack(data.read(2))[0]
    else:
        value = STRUCT_UINT16.unpack(data[index:index+2])[0]
        index += 2
    return (value, index)


def read_length_prefixed_string(data, index):
    """
    Read a length-prefixed ANSI string from data.

    :param data: bytes string or stream containing the data to be extracted.
    :param index: index in data where string size start or None if data is
                  stream
    :return: tuple (value, index) containing the read value (bytes string),
             and the index to start reading from next time.
    """
    length, index = read_uint32(data, index)
    # if length = 0, return a null string (no null character)
    if length == 0:
        return ('', index)
    # extract the string without the last null character
    if index is None:
        ansi_string = data.read(length-1)
        null_char = data.read(1)
    else:
        ansi_string = data[index:index+length-1]
        null_char = data[index+length-1]
        index += length
    # TODO: only in strict mode:
    # check the presence of the null char:
    assert null_char == NULL_CHAR
    return (ansi_string, index)


def guess_encoding(data):
    """ guess encoding of byte string to create unicode

    Since this is used to decode path names from ole objects, prefer latin1
    over utf* codecs if ascii is not enough
    """
    for encoding in 'ascii', 'latin1', 'utf8', 'utf-16-le', 'utf16':
        try:
            result = data.decode(encoding, errors='strict')
            log.debug(u'decoded using {0}: "{1}"'.format(encoding, result))
            return result
        except UnicodeError:
            pass
    log.warning('failed to guess encoding for string, falling back to '
                'ascii with replace')
    return data.decode('ascii', errors='replace')


def read_zero_terminated_string(data, index):
    """
    Read a zero-terminated string from data

    :param data: bytes string or stream containing an ansi string
    :param index: index at which the string should start or None if data is
                  stream
    :return: tuple (unicode, index) containing the read string (unicode),
             and the index to start reading from next time.
    """
    if index is None:
        result = bytearray()
        for _ in xrange(STR_MAX_LEN):
            char = ord(data.read(1))    # need ord() for py3
            if char == 0:
                return guess_encoding(result), index
            result.append(char)
        raise ValueError('found no string-terminating zero-byte!')
    else:       # data is byte array, can just search
        end_idx = data.index(b'\x00', index, index+STR_MAX_LEN)
        # encode and return with index after the 0-byte
        return guess_encoding(data[index:end_idx]), end_idx+1


# === CLASSES =================================================================


class OleNativeStream(object):
    """
    OLE object contained into an OLENativeStream structure.
    (see MS-OLEDS 2.3.6 OLENativeStream)

    Filename and paths are decoded to unicode.
    """
    # constants for the type attribute:
    # see MS-OLEDS 2.2.4 ObjectHeader
    TYPE_LINKED = 0x01
    TYPE_EMBEDDED = 0x02

    def __init__(self, bindata=None, package=False):
        """
        Constructor for OleNativeStream.
        If bindata is provided, it will be parsed using the parse() method.

        :param bindata: forwarded to parse, see docu there
        :param package: bool, set to True when extracting from an OLE Package
                        object
        """
        self.filename = None
        self.src_path = None
        self.unknown_short = None
        self.unknown_long_1 = None
        self.unknown_long_2 = None
        self.temp_path = None
        self.actual_size = None
        self.data = None
        self.package = package
        self.is_link = None
        self.data_is_stream = None
        if bindata is not None:
            self.parse(data=bindata)

    def parse(self, data):
        """
        Parse binary data containing an OLENativeStream structure,
        to extract the OLE object it contains.
        (see MS-OLEDS 2.3.6 OLENativeStream)

        :param data: bytes array or stream, containing OLENativeStream
                     structure containing an OLE object
        :return: None
        """
        # TODO: strict mode to raise exceptions when values are incorrect
        # (permissive mode by default)
        if hasattr(data, 'read'):
            self.data_is_stream = True
            index = None       # marker for read_* functions to expect stream
        else:
            self.data_is_stream = False
            index = 0          # marker for read_* functions to expect array

        # An OLE Package object does not have the native data size field
        if not self.package:
            self.native_data_size, index = read_uint32(data, index)
            log.debug('OLE native data size = {0:08X} ({0} bytes)'
                      .format(self.native_data_size))
        # I thought this might be an OLE type specifier ???
        self.unknown_short, index = read_uint16(data, index)
        self.filename, index = read_zero_terminated_string(data, index)
        # source path
        self.src_path, index = read_zero_terminated_string(data, index)
        # TODO: I bet these 8 bytes are a timestamp ==> FILETIME from olefile
        self.unknown_long_1, index = read_uint32(data, index)
        self.unknown_long_2, index = read_uint32(data, index)
        # temp path?
        self.temp_path, index = read_zero_terminated_string(data, index)
        # size of the rest of the data
        try:
            self.actual_size, index = read_uint32(data, index)
            if self.data_is_stream:
                self.data = data
            else:
                self.data = data[index:index+self.actual_size]
            self.is_link = False
            # TODO: there can be extra data, no idea what it is for
            # TODO: SLACK DATA
        except (IOError, struct.error):      # no data to read actual_size
            log.debug('data is not embedded but only a link')
            self.is_link = True
            self.actual_size = 0
            self.data = None


class OleObject(object):
    """
    OLE 1.0 Object

    see MS-OLEDS 2.2 OLE1.0 Format Structures
    """

    # constants for the format_id attribute:
    # see MS-OLEDS 2.2.4 ObjectHeader
    TYPE_LINKED = 0x01
    TYPE_EMBEDDED = 0x02

    def __init__(self, bindata=None):
        """
        Constructor for OleObject.
        If bindata is provided, it will be parsed using the parse() method.

        :param bindata: bytes, OLE 1.0 Object structure containing OLE object

        Note: Code can easily by generalized to work with byte streams instead
              of arrays just like in OleNativeStream.
        """
        self.ole_version = None
        self.format_id = None
        self.class_name = None
        self.topic_name = None
        self.item_name = None
        self.data = None
        self.data_size = None
        if bindata is not None:
            self.parse(bindata)

    def parse(self, data):
        """
        Parse binary data containing an OLE 1.0 Object structure,
        to extract the OLE object it contains.
        (see MS-OLEDS 2.2 OLE1.0 Format Structures)

        :param data: bytes, OLE 1.0 Object structure containing an OLE object
        :return:
        """
        # from ezhexviewer import hexdump3
        # print("Parsing OLE object data:")
        # print(hexdump3(data, length=16))
        # Header: see MS-OLEDS 2.2.4 ObjectHeader
        index = 0
        self.ole_version, index = read_uint32(data, index)
        self.format_id, index = read_uint32(data, index)
        log.debug('OLE version=%08X - Format ID=%08X',
                  self.ole_version, self.format_id)
        assert self.format_id in (self.TYPE_EMBEDDED, self.TYPE_LINKED)
        self.class_name, index = read_length_prefixed_string(data, index)
        self.topic_name, index = read_length_prefixed_string(data, index)
        self.item_name, index = read_length_prefixed_string(data, index)
        log.debug('Class name=%r - Topic name=%r - Item name=%r',
                  self.class_name, self.topic_name, self.item_name)
        if self.format_id == self.TYPE_EMBEDDED:
            # Embedded object: see MS-OLEDS 2.2.5 EmbeddedObject
            # assert self.topic_name != '' and self.item_name != ''
            self.data_size, index = read_uint32(data, index)
            log.debug('Declared data size=%d - remaining size=%d',
                      self.data_size, len(data)-index)
            # TODO: handle incorrect size to avoid exception
            self.data = data[index:index+self.data_size]
            assert len(self.data) == self.data_size
            self.extra_data = data[index+self.data_size:]


def shorten_filename(fname, max_len):
    """Create filename shorter than max_len, trying to preserve suffix."""
    # simple cases:
    if not max_len:
        return fname
    name_len = len(fname)
    if name_len < max_len:
        return fname

    idx = fname.rfind('.')
    if idx == -1:
        return fname[:max_len]

    suffix_len = name_len - idx  # length of suffix including '.'
    if suffix_len > max_len:
        return fname[:max_len]

    # great, can preserve suffix
    return fname[:max_len-suffix_len] + fname[idx:]


def sanitize_filename(filename, replacement='_',
                      max_len=MAX_FILENAME_LENGTH):
    """
    Return filename that is save to work with.

    Removes path components, replaces all non-whitelisted characters (so output
    is always a pure-ascii string), replaces '..' and '  ' and shortens to
    given max length, trying to preserve suffix.

    Might return empty string
    """
    basepath = os.path.basename(filename).strip()
    sane_fname = re.sub(u'[^a-zA-Z0-9.\-_ ]', replacement, basepath)
    sane_fname = str(sane_fname)    # py3: does nothing;   py2: unicode --> str

    while ".." in sane_fname:
        sane_fname = sane_fname.replace('..', '.')

    while "  " in sane_fname:
        sane_fname = sane_fname.replace('  ', ' ')

    # limit filename length, try to preserve suffix
    return shorten_filename(sane_fname, max_len)


def get_sane_embedded_filenames(filename, src_path, tmp_path, max_len,
                                noname_index):
    """
    Get some sane filenames out of path information, preserving file suffix.

    Returns several canddiates, first with suffix, then without, then random
    with suffix and finally one last attempt ignoring max_len using arg
    `noname_index`.

    In some malware examples, filename (on which we relied sofar exclusively
    for this) is empty or " ", but src_path and tmp_path contain paths with
    proper file names. Try to extract filename from any of those.

    Preservation of suffix is especially important since that controls how
    windoze treats the file.
    """
    suffixes = []
    candidates_without_suffix = []  # remember these as fallback
    for candidate in (filename, src_path, tmp_path):
        # remove path component. Could be from linux, mac or windows
        idx = max(candidate.rfind('/'), candidate.rfind('\\'))
        candidate = candidate[idx+1:].strip()

        # sanitize
        candidate = sanitize_filename(candidate, max_len=max_len)

        if not candidate:
            continue    # skip whitespace-only

        # identify suffix. Dangerous suffixes are all short
        idx = candidate.rfind('.')
        if idx is -1:
            candidates_without_suffix.append(candidate)
            continue
        elif idx < len(candidate)-5:
            candidates_without_suffix.append(candidate)
            continue

        # remember suffix
        suffixes.append(candidate[idx:])

        yield candidate

    # parts with suffix not good enough? try those without one
    for candidate in candidates_without_suffix:
        yield candidate

    # then try random
    suffixes.append('')  # ensure there is something in there
    for _ in range(MAX_FILENAME_ATTEMPTS):
        for suffix in suffixes:
            leftover_len = max_len - len(suffix)
            if leftover_len < 1:
                continue
            name = ''.join(random.sample('abcdefghijklmnopqrstuvwxyz',
                                         min(26, leftover_len)))
            yield name + suffix

    # still not returned? Then we have to make up a name ourselves
    # do not care any more about max_len (maybe it was 0 or negative)
    yield 'oleobj_%03d' % noname_index


def find_ole_in_ppt(filename):
    """ find ole streams in ppt

    This may be a bit confusing: we get an ole file (or its name) as input and
    as output we produce possibly several ole files. This is because the
    data structure can be pretty nested:
    A ppt file has many streams that consist of records. Some of these records
    can contain data which contains data for another complete ole file (which
    we yield). This embedded ole file can have several streams, one of which
    can contain the actual embedded file we are looking for (caller will check
    for these).
    """
    ppt_file = None
    try:
        ppt_file = PptFile(filename)
        for stream in ppt_file.iter_streams():
            for record_idx, record in enumerate(stream.iter_records()):
                if isinstance(record, PptRecordExOleVbaActiveXAtom):
                    ole = None
                    try:
                        data_start = next(record.iter_uncompressed())
                        if data_start[:len(olefile.MAGIC)] != olefile.MAGIC:
                            continue   # could be ActiveX control / VBA Storage

                        # otherwise, this should be an OLE object
                        log.debug('Found record with embedded ole object in '
                                  'ppt (stream "{0}", record no {1})'
                                  .format(stream.name, record_idx))
                        ole = record.get_data_as_olefile()
                        yield ole
                    except IOError:
                        log.warning('Error reading data from {0} stream or '
                                    'interpreting it as OLE object'
                                    .format(stream.name))
                        log.debug('', exc_info=True)
                    finally:
                        if ole is not None:
                            ole.close()
    finally:
        if ppt_file is not None:
            ppt_file.close()


class FakeFile(io.RawIOBase):
    """ 从数据创建类似文件的对象而不复制

    BytesIO是我想要使用的，但它会复制所有数据。
    这个类没有。不利的一面是：数据只能读取和查找，而不能写入。

    假设给定的数据是字节（在py2中为str，在py3中为字节）。

    另请参阅（也许可以与一起放入公共文件）：
    ppt_record_parser.IterStream，ooxml.ZipSubFile
    """

    def __init__(self, data):
        """ 使用给定字节的数据创建FakeFile """
        super(FakeFile, self).__init__()
        self.data = data   # 这实际上并没有复制（python很懒惰）
        self.pos = 0
        self.size = len(data)

    def readable(self):
        return True

    def writable(self):
        return False

    def seekable(self):
        return True

    def readinto(self, target):
        """ 读入预先分配的目标 """
        n_data = min(len(target), self.size-self.pos)
        if n_data == 0:
            return 0
        target[:n_data] = self.data[self.pos:self.pos+n_data]
        self.pos += n_data
        return n_data

    def read(self, n_data=-1):
        """ 读取并返回数据 """
        if self.pos >= self.size:
            return bytes()
        if n_data == -1:
            n_data = self.size - self.pos
        result = self.data[self.pos:self.pos+n_data]
        self.pos += n_data
        return result

    def seek(self, pos, offset=io.SEEK_SET):
        """ 跳转到文件中的另一个位置 """
        # 根据self-pos、pos和offset计算目标位置
        if offset == io.SEEK_SET:
            new_pos = pos
        elif offset == io.SEEK_CUR:
            new_pos = self.pos + pos
        elif offset == io.SEEK_END:
            new_pos = self.size + pos
        else:
            raise ValueError("偏移量｛0｝无效，需要SEEK_*常量"
                             .format(offset))
        if new_pos < 0:
            raise IOError('不允许在文件开头以外搜索')
        self.pos = new_pos

    def tell(self):
        """ 告诉我们在文件中的位置 """
        return self.pos


def find_ole(filename, data, xml_parser=None):
    """ 尝试打开 zip/ole/rtf/... ; yield None 如果失败

    如果给定了数据，则文件名（一般情况下）被忽略。

    以OleFileIO的形式返回嵌入的ole流。
    """

    if data is not None:
        # isOleFile和is_ppt可以直接处理数据，但zip需要文件
        # --> 将数据包装在类似文件的对象中而不复制数据
        log.debug('处理数据，下面的文件未被触及')
        arg_for_ole = data
        arg_for_zip = FakeFile(data)
    else:
        # 我们只有一个文件名
        log.debug('按名称处理文件')
        arg_for_ole = filename
        arg_for_zip = filename

    ole = None
    try:
        if olefile.isOleFile(arg_for_ole):
            if is_ppt(arg_for_ole):
                log.info('是PPT文件： ' + filename)
                for ole in find_ole_in_ppt(arg_for_ole):
                    yield ole
                    ole = None   # 在 find_ole_in_ppt 中关闭
            # 无论如何：检查非扇区流中的嵌入内容
            log.info('是OLE文件： ' + filename)
            ole = olefile.OleFileIO(arg_for_ole)
            yield ole
        elif xml_parser is not None or is_zipfile(arg_for_zip):
            # 与调用此函数的第三方代码保持兼容性
            # 直接执行，而不提供XmlParser实例
            if xml_parser is None:
                xml_parser = XmlParser(arg_for_zip)
                # 强制迭代，使 XmlParser.iter_no_xml() 返回数据
                for _ in xml_parser.iter_xml():
                    pass

            log.info('是zip文件: ' + filename)
            # 我们之前遍历了XML文件，
            # 现在我们可以迭代非XML文件来查找ole对象
            for subfile, _, file_handle in xml_parser.iter_non_xml():
                try:
                    head = file_handle.read(len(olefile.MAGIC))
                except RuntimeError:
                    log.error('zip已加密： ' + filename)
                    yield None
                    continue

                if head == olefile.MAGIC:
                    file_handle.seek(0)
                    log.info('解压缩ole： ' + subfile)
                    try:
                        ole = olefile.OleFileIO(file_handle)
                        yield ole
                    except IOError:
                        log.warning('从｛0｝/｛1｝读取数据或'
                                    '将其解释为OLE对象时出错'
                                    .format(filename, subfile))
                        log.debug('', exc_info=True)
                    finally:
                        if ole is not None:
                            ole.close()
                            ole = None
                else:
                    log.debug('跳过解压: ' + subfile)
        else:
            log.warning('打开文件失败: {0} (或者它是数据) 既不是zip也不是OLE'
                        .format(filename))
            yield None
    except Exception:
        log.error('打开｛0｝时出现异常'.format(filename),
                  exc_info=True)
        yield None
    finally:
        if ole is not None:
            ole.close()


def find_external_relationships(xml_parser):
    """ iterate XML files looking for relationships to external objects
    """
    for _, elem, _ in xml_parser.iter_xml(None, False, OOXML_RELATIONSHIP_TAG):
        try:
            if elem.attrib['TargetMode'] == 'External':
                relationship_type = elem.attrib['Type'].rsplit('/', 1)[1]

                if relationship_type in BLACKLISTED_RELATIONSHIP_TYPES:
                    yield relationship_type, elem.attrib['Target']
        except (AttributeError, KeyError):
            # ignore missing attributes - Word won't detect
            # external links anyway
            pass


def find_customUI(xml_parser):
    """
    iterate XML files looking for customUI to external objects or VBA macros
    Examples of malicious usage, to load an external document or trigger a VBA macro:
    https://www.trellix.com/en-us/about/newsroom/stories/threat-labs/prime-ministers-office-compromised.html
    https://www.netero1010-securitylab.com/evasion/execution-of-remote-vba-script-in-excel
    """
    for _, elem, _ in xml_parser.iter_xml(None, False, (TAG_CUSTOMUI_2007, TAG_CUSTOMUI_2010)):
       customui_onload = elem.get('onLoad')
       if customui_onload is not None:
            yield customui_onload


def process_file(filename, data, output_dir=None):
    """ 在给定文件中查找嵌入对象

    如果给定了data（来自加密zip文件的xglob），
    则不用filename读取文件。否则（一般），
    则根据需要从filename中读取数据。

    如果给了output_dir，但不存在，则创建它。
    否则将数据保存到与输入文件相同的目录中。
    """
    # 清除文件名，为嵌入的文件名部分留出空间
    sane_fname = sanitize_filename(filename, max_len=MAX_FILENAME_LENGTH-5) or\
        'NONAME'
    if output_dir:
        if not os.path.isdir(output_dir):
            log.info('创建输出目录 %s', output_dir)
            os.mkdir(output_dir)

        fname_prefix = os.path.join(output_dir, sane_fname)
    else:
        base_dir = os.path.dirname(filename)
        fname_prefix = os.path.join(base_dir, sane_fname)

    # TODO: 将对象提取到文件的选项（默认情况下为 false）
    print('-'*79)
    print('文件： %r' % filename)
    index = 1

    # 不要抛出错误，但要记住它们，并尝试继续使用其他流
    err_stream = False
    err_dumping = False
    did_dump = False

    xml_parser = None
    if is_zipfile(filename):
        log.info('文件可以是OOXML文件，用于查找与'
                 '外部链接的关系')
        xml_parser = XmlParser(filename)
        for relationship, target in find_external_relationships(xml_parser):
            did_dump = True
            print("找到了与外部链接 %s 的关系 '%s'" % (relationship, target))
            if target.startswith('mhtml:'):
                print("潜在的漏洞： CVE-2021-40444")
        for target in find_customUI(xml_parser):
            did_dump = True
            print("发现具有外部链接或VBA宏 %s 的自定义用户界面标记（可能利用CVE-2021-42292）" % target)

    # 在文件中查找 ole 文件 (例如 unzip docx)
    # 必须在迭代中完成每个ole流的工作，
    # 因为句柄在find_ole中是关闭的
    for ole in find_ole(filename, data, xml_parser):
        if ole is None:    # 未找到 ole 对象
            continue

        for path_parts in ole.listdir():
            stream_path = '/'.join(path_parts)
            log.debug('正在检查流：%r', stream_path)
            if path_parts[-1].lower() == '\x01ole10native':
                stream = None
                try:
                    stream = ole.openstream(path_parts)
                    print('从流 %r 中提取嵌入OLE对象中的文件：'
                          % stream_path)
                    print('正在分析OLE包')
                    opkg = OleNativeStream(stream)
                    # 让数据流保持畅通，直到转存完成
                except Exception:
                    log.warning('*** 不是一个 OLE 1.0 对象')
                    err_stream = True
                    if stream is not None:
                        stream.close()
                    continue

                # 打印信息
                if opkg.is_link:
                    log.debug('对象是链接，而不是嵌入的文件 '
                              '- 跳过')
                    continue
                print(u'文件名 = "%s"' % opkg.filename)
                print(u'原路径 = "%s"' % opkg.src_path)
                print(u'缓存路径 = "%s"' % opkg.temp_path)
                for embedded_fname in get_sane_embedded_filenames(
                        opkg.filename, opkg.src_path, opkg.temp_path,
                        MAX_FILENAME_LENGTH - len(sane_fname) - 1, index):
                    fname = fname_prefix + '_' + embedded_fname
                    if not os.path.isfile(fname):
                        break

                # 转存
                try:
                    print('正在保存到文件中： %s' % fname)
                    with open(fname, 'wb') as writer:
                        n_dumped = 0
                        next_size = min(DUMP_CHUNK_SIZE, opkg.actual_size)
                        while next_size:
                            data = stream.read(next_size)
                            writer.write(data)
                            n_dumped += len(data)
                            if len(data) != next_size:
                                log.warning('想要读取： {0}, 实际读取： {1}'
                                            .format(next_size, len(data)))
                                break
                            next_size = min(DUMP_CHUNK_SIZE,
                                            opkg.actual_size - n_dumped)
                    did_dump = True
                except Exception as exc:
                    log.warning('转存错误： {0} ({1})'
                                .format(fname, exc))
                    err_dumping = True
                finally:
                    stream.close()

                index += 1
    return err_stream, err_dumping, did_dump


# === 主要的 ====================================================================


def existing_file(filename):
    """ 由参数分析器调用以查看给定文件是否存在 """
    if not os.path.isfile(filename):
        raise argparse.ArgumentTypeError('{0} is not a file.'.format(filename))
    return filename


def main(cmd_line_args=None):
    """ 主函数，作为脚本运行时调用

    默认情况下（cmd_line_args=None）使用sys.argv。但是，对于测试，可以
    提供其他参数。
    """
    # 打印带有版本的横幅
    ensure_stdout_handles_unicode()
    print('oleobj %s - http://decalage.info/oletools' % __version__)
    print('这是正在进行的工作 - 定期检查更新！')
    print('如有任何问题，请访问 '
          'https://github.com/decalage2/oletools/issues')
    print('')

    usage = '用法: %(prog)s [选项] <文件名> [文件名2 ...]'
    parser = argparse.ArgumentParser(usage=usage)
    # parser.add_argument('-o', '--outfile', dest='outfile',
    #     help='output file')
    # parser.add_argument('-c', '--csv', dest='csv',
    #     help='export results to a CSV file')
    parser.add_argument("-r", action="store_true", dest="recursive",
                        help='在子目录中递归查找文件。')
    parser.add_argument("-d", type=str, dest="output_dir", default=None,
                        help='使用指定的目录输出文件。')
    parser.add_argument("-z", "--zip", dest='zip_password', type=str,
                        default=None,
                        help='如果文件是zip档案，使用'
                             '提供的密码打开其中的第一'
                             '个文件（需要Python2.6+）')
    parser.add_argument("-f", "--zipfname", dest='zip_fname', type=str,
                        default='*',
                        help='如果文件是zip档案，则表示'
                             '要在zip中打开的文件。'
                             '支持通配符*和？。（默认值：*）')
    parser.add_argument('-l', '--loglevel', dest="loglevel", action="store",
                        default=DEFAULT_LOG_LEVEL,
                        help='日志记录级别debug/info/warning/error/critical'
                             '（默认值=%(default)s）')
    parser.add_argument('input', nargs='*', type=existing_file, metavar='FILE',
                        help='要分析的Office文件（与-i相同）')

    # 与ripOLE兼容的选项
    parser.add_argument('-i', '--more-input', type=str, metavar='FILE',
                        help='要分析的附加文件'
                             '（与位置参数相同）')
    parser.add_argument('-v', '--verbose', action='store_true',
                        help='详细模式，将日志记录设置为DEBUG'
                             '（重写 -l）')

    options = parser.parse_args(cmd_line_args)
    if options.more_input:
        options.input += [options.more_input, ]
    if options.verbose:
        options.loglevel = 'debug'

    # 如果未传递参数，则打印帮助信息
    if not options.input:
        parser.print_help()
        return RETURN_ERR_ARGS

    # 设置控制台日志：
    # 在这里，我们默认使用stdout而不是stderr，这样
    # 可以正确地重定向输出。
    logging.basicConfig(level=LOG_LEVELS[options.loglevel], stream=sys.stdout,
                        format='%(levelname)-8s %(message)s')
    # 启用日志模块：
    log.setLevel(logging.NOTSET)
    if options.loglevel == 'debug-olefile':
        olefile.enable_logging()

    # 记住是否存在问题，然后继续处理其他数据
    any_err_stream = False
    any_err_dumping = False
    any_did_dump = False

    for container, filename, data in \
            xglob.iter_files(options.input, recursive=options.recursive,
                             zip_password=options.zip_password,
                             zip_fname=options.zip_fname):
        # 忽略存储在zip文件中的目录名：
        if container and filename.endswith('/'):
            continue
        err_stream, err_dumping, did_dump = \
            process_file(filename, data, options.output_dir)
        any_err_stream |= err_stream
        any_err_dumping |= err_dumping
        any_did_dump |= did_dump

    # 总结返回值
    return_val = RETURN_NO_DUMP
    if any_did_dump:
        return_val += RETURN_DID_DUMP
    if any_err_stream:
        return_val += RETURN_ERR_STREAM
    if any_err_dumping:
        return_val += RETURN_ERR_DUMP
    return return_val


if __name__ == '__main__':
    sys.exit(main())
