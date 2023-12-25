import platform
revar1 = platform.version()

if platform.system() == 'Windows':
    if platform.version().startswith('10'):
        if platform.version().replace('.', '').isdigit() and int(platform.version().replace('.', '')) >= 10022000:
            revar1 = 'Windows11'
        else:
            revar1 = 'Windows10'
    else:
        if platform.release() == '10':
            revar1 = 'Windows10'
        elif platform.release() == '8':
            revar1 = 'Windows8'
        elif platform.release() == '7':
            revar1 = 'Windows7'
        else:
            revar1 = 'OtherWindows'
else:
    revar1 = 'NotWindows'

print(revar1)