from cx_Freeze import setup, Executable

buildOptions = {
    'include_files': [
        'assets/'
    ]
}

base = 'Win32GUI'

executables = [
    Executable('MailBot.py', base=base, icon='assets/app.ico'),
]

setup(
    name='MailBot',
    version='0.1',
    description='Dark Powered MailBot',
    options=dict(build_exe=buildOptions),
    executables=executables
)