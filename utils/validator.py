import ipaddress
from typing import override

from qfluentwidgets import ConfigValidator


class IPValidator(ConfigValidator):
    """IP validator"""

    @override
    def validate(self, value) -> bool:
        try:
            ipaddress.ip_address(value)
        except ValueError:
            return False
        return True

    @override
    def correct(self, value):
        return "127.0.0.1"
