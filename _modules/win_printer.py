# -*- coding: utf-8 -*-
'''
Windows Printer Port Management

:platform:      Windows

'''


from __future__ import absolute_import

# Import salt libs
import salt.utils

# Define the module's virtual name
__virtualname__ = 'win_printer'


def __virtual__():
    '''
    Load only on Windows
    '''
    if salt.utils.is_windows():
        return __virtualname__
    return (False, 'Module win_printer: module only works on Windows systems')


def _srvmgr(func):
    '''
    Execute a function from the WebAdministration PS module
    '''

    return __salt__['cmd.run'](
        'Import-Module WebAdministration; {0}'.format(func),
        shell='powershell',
        python_shell=True)


def list_printers():
    '''
    List all installed ports

    CLI Example:

    .. code-block:: bash

        salt '*' win_printer.list_printers
    '''
    pscmd = []
    pscmd.append(r'Get-WebSite -erroraction silentlycontinue -warningaction silentlycontinue')
    pscmd.append(r' | foreach {')
    pscmd.append(r' $_.Name')
    pscmd.append(r'};')

    command = ''.join(pscmd)
    return _srvmgr(command)


def create_printer(
        name,
        protocol,
        sourcepath,
        port,
        apppool='',
        hostheader='',
        ipaddress=''):
    '''
    Create a basic webport in IIS

    CLI Example:

    .. code-block:: bash

        salt '*' win_printer.create_printer name='My Test Site' protocol='http' sourcepath='c:\\stage' port='80' apppool='TestPool'

    '''

    pscmd = []
    pscmd.append(r'cd IIS:\Sites\;')
    pscmd.append(r'New-Item \'iis:\Sites\{0}\''.format(name))
    pscmd.append(r' -bindings @{{protocol=\'{0}\';bindingInformation=\':{1}:{2}\'}}'.format(
        protocol, port, hostheader.replace(' ', '')))
    pscmd.append(r'-physicalPath {0};'.format(sourcepath))

    if apppool:
        pscmd.append(r'Set-ItemProperty \'iis:\Sites\{0}\''.format(name))
        pscmd.append(r' -name applicationPool -value \'{0}\';'.format(apppool))

    command = ''.join(pscmd)
    return _srvmgr(command)


def remove_printer(name):
    '''
    Delete printer port

    CLI Example:

    .. code-block:: bash

        salt '*' win_printer.remove_printer name='My Test Site'

    '''

    pscmd = []
    pscmd.append(r'cd IIS:\Sites\;')
    pscmd.append(r'Remove-WebSite -Name \'{0}\''.format(name))

    command = ''.join(pscmd)
    return _srvmgr(command)


def update_printer(name, **todo):
    '''
    Update a printer port
    '''

    return 'TODO'
