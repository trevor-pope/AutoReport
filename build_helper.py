"""
Generates a list of files and directories to include when building the installer using NSIS.
"""
import os


if __name__ == '__main__':
    print('Scanning for files to install')

    root = 'dist\\autoreport'
    stack = []
    install_contents = ['# Files and dirs to install\n']

    for cur_dir, folders, files in os.walk('dist\\autoreport', topdown=True):
        if cur_dir != root:
            install_contents.append(f'SetOutPath "$INSTDIR\\{cur_dir[len(root):]}"\n')
            stack.append(('dir', cur_dir[len(root) + 1:]))

        for file in files:
            install_contents.append(f'File "{os.path.join(cur_dir, file)}"\n')
            stack.append(('file', os.path.join(cur_dir[len(root) + 1:], file)))

    print('Writing file paths to install_list.nsh')
    with open('install_list.nsh', 'w') as f:
        f.writelines(install_contents)


    uninstall_contents = ['# Files and dirs to uninstall\n']

    for kind, name in reversed(stack):
        if kind == 'dir':
            uninstall_contents.append(f'RMDir "$INSTDIR\\{name}"\n')
        else:
            uninstall_contents.append(f'Delete "$INSTDIR\\{name}"\n')

    print('Writing file paths in reverse to uninstall_list.nsh')
    with open('uninstall_list.nsh', 'w') as f:
        f.writelines(uninstall_contents)

    print('Done!')
