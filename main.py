# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import platform


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    SYSTEM_PLATFORM = ''
    if platform.system() == 'Windows':
        if platform.version().startswith('10'):
            if platform.version().replace('.', '').isdigit() and int(platform.version().replace('.', '')) >= 10022000:
                #print("Windows 11")
                SYSTEM_PLATFORM = 'Windows11'
            else:
                print("Not Windows")
        else:
            if platform.release() == '10':
                #print("Windows 10")
                SYSTEM_PLATFORM = 'Windows10'
            elif platform.release() == '8':
                SYSTEM_PLATFORM = 'Windows8'
            elif platform.release() == '7':
                SYSTEM_PLATFORM = 'Windows7'
            else:
                #print("Other Windows")
                SYSTEM_PLATFORM = 'Other Windows'
    else:
        print("Not Windows")
    print(SYSTEM_PLATFORM)
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
