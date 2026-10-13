from dataclasses import dataclass

THEMES = {
    'nankai': dict(primary='7E0C6E', deep='43243E', soft='F3EDF1', accent='AD879F'),
    'slate': dict(primary='27465E', deep='1B3041', soft='EAF0F4', accent='7595AA'),
    'wine': dict(primary='7A2B22', deep='432B2B', soft='F5EFEB', accent='BD9288'),
}

@dataclass
class Theme:
    primary: str = '7E0C6E'
    deep: str = '43243E'
    soft: str = 'F3EDF1'
    accent: str = 'AD879F'
    ink: str = '242424'
    body: str = '404040'
    muted: str = '69666A'
    line: str = 'E6E0E5'
    white: str = 'FFFFFF'
    font_cn: str = 'Noto Sans CJK SC'
    font_en: str = 'Aptos'

    @classmethod
    def from_spec(cls, spec):
        if isinstance(spec, str):
            vals = THEMES.get(spec)
            if vals is None: raise ValueError(f'Unknown theme: {spec}; use {list(THEMES)}')
        elif isinstance(spec, dict): vals = spec
        else: vals = THEMES['nankai']
        obj = cls()
        for k,v in vals.items():
            if not hasattr(obj, k): raise ValueError(f'Unknown theme property {k}')
            if k not in ('font_cn', 'font_en'):
                v = v.removeprefix('#').upper()
                if len(v) != 6 or any(c not in '0123456789ABCDEF' for c in v):
                    raise ValueError(f'Invalid hex colour: {k}={v}')
            setattr(obj, k, v)
        return obj
