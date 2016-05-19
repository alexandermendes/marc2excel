# -*- coding: utf-8 -*-
"""Exceptions module for marc2excel."""


class InvalidMARCCodeError(ValueError):
    """An invalid MARC code was found.

    :param code: The invalid MARC code.
    :param path: The file in which the code was found.
    """

    def __init__(self, code, path):
        msg = '"{0}" ({1})'.format(code, path)
        super(InvalidMARCCodeError, self).__init__(msg)
