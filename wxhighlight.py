#!/usr/bin/env python
################################################################################
#
# \file
# \author   <a href="http://www.innomatic.ca">innomatic</a>
# \brief    Highlight wxPython frontend
# \see      http://www.andre-simon.de/doku/highlight/en/highlight.php
# \bug      screen capture works on the first monitor only.
#

import ctypes
import glob
import os
import pickle
import wx
import wx.html2 as html2
from subprocess import run, PIPE

## choice box size limit
CTRL_SIZE=(140,24)

## Get the real screen size after SetProcessDPIAware, which is used to compute
#  the screen ratio. This function runs in separate process not to affect the UI.
def SystemMetrics(q):
    ctypes.windll.user32.SetProcessDPIAware()
    q.put(ctypes.windll.user32.GetSystemMetrics(0))
    return


## FileDropTarget. On drop it loads the file contents and update the file name
class MyFileDropTarget(wx.FileDropTarget):

    def __init__(self, window, log):
        wx.FileDropTarget.__init__(self)
        # target window
        self.window = window
        # file name window (wx.TextCtrl)
        self.log = log

    # Load file content into the target window
    def OnDropFiles(self, x, y, fname):
        try:
            # let the parent load the file
            self.window.LoadFile(fname[0])
        except:
            return False
        else:
            # show file name on the log window
            if 'nt' in os.name:
                self.log.SetValue(fname[0][fname[0].rfind('\\')+1:])
            else:
                self.log.SetValue(fname[0][fname[0].rfind('/')+1:])

            return True


## Main frame window
class MyFrame(wx.Frame):

    def __init__(self, *args, **kwgs):
        wx.Frame.__init__(self, *args, **kwgs)

        # spliter window 
        self.spw = wx.SplitterWindow(self, style=wx.SP_LIVE_UPDATE)
        self.spw.SetMinimumPaneSize(100)

        # source text on the top
        self.textSrc = wx.TextCtrl(self.spw, style=wx.TE_MULTILINE)

        # notebook below
        self.nbkOut = wx.Notebook(self.spw)
        self.spw.SplitHorizontally(self.textSrc, self.nbkOut, 0)

        # webview output
        self.webView = html2.WebView.New(self.nbkOut)
        # it becomes the first page of the notebook
        self.nbkOut.AddPage(self.webView, 'WebView')

        # text output
        self.textOut = wx.TextCtrl(self.nbkOut,
                style=wx.TE_MULTILINE|wx.TE_READONLY)
        # it becomes the second page of the notebook
        self.nbkOut.AddPage(self.textOut, 'TextView')

        # control panel on the right
        self.pnlCtrl = wx.Panel(self)

        # file name
        self.sttFlname = wx.StaticText(self.pnlCtrl, -1, 'File')
        self.txtFlname = wx.TextCtrl(self.pnlCtrl, -1, style=wx.TE_READONLY)
        # display scale
        self.sttDisply = wx.StaticText(self.pnlCtrl, -1, 'Scale')
        self.txtDisply = wx.TextCtrl(self.pnlCtrl, -1, style=wx.TE_READONLY)
        # txtFlname events
        self.txtFlname.Bind(wx.EVT_LEFT_DOWN, self.OnLoadSource)
        self.txtFlname.Bind(wx.EVT_TEXT, self.OnSourceName)

        # other controls
        self.sttSyntax = wx.StaticText(self.pnlCtrl, -1, 'Syntax')
        self.choSyntax = wx.Choice(self.pnlCtrl, -1, style=wx.CB_SORT)
        self.sttOutput = wx.StaticText(self.pnlCtrl, -1, 'Output')
        self.choOutput = wx.Choice(self.pnlCtrl, -1, style=wx.CB_SORT)
        self.sttThemes = wx.StaticText(self.pnlCtrl, -1, 'Theme')
        self.choThemes = wx.Choice(self.pnlCtrl, -1, style=wx.CB_SORT)
        self.sttAstyle = wx.StaticText(self.pnlCtrl, -1, 'Astyle')
        self.choAstyle = wx.Choice(self.pnlCtrl, -1, style=wx.CB_SORT)
        self.sttHlFont = wx.StaticText(self.pnlCtrl, -1, 'Font')
        self.choHlFont = wx.Choice(self.pnlCtrl, -1, style=wx.CB_SORT)
        self.sttFntSiz = wx.StaticText(self.pnlCtrl, -1, 'Size')
        self.choFntSiz = wx.Choice(self.pnlCtrl, -1)

        # limit the size of the choice boxes
        self.choSyntax.SetMaxSize(CTRL_SIZE)
        self.choOutput.SetMaxSize(CTRL_SIZE)
        self.choThemes.SetMaxSize(CTRL_SIZE)

        # option checkboxes
        self.sttOption = wx.StaticText(self.pnlCtrl, -1, 'Option')
        self.chkLineNo = wx.CheckBox(self.pnlCtrl, -1, 'Line Numbering')
        self.chkWrapLn = wx.CheckBox(self.pnlCtrl, -1, 'Wrap Lines after 80')
        self.chkInLCss = wx.CheckBox(self.pnlCtrl, -1, 'CSS within each tag')

        # convert button
        self.btnConvrt = wx.Button(self.pnlCtrl, -1, label='Convert')
        self.Bind(wx.EVT_BUTTON, self.OnConvert, self.btnConvrt)
        # text to clipboard
        self.btnClpTxt = wx.Button(self.pnlCtrl, -1, label='Output to Clipboard')
        self.Bind(wx.EVT_BUTTON, self.OnClipText, self.btnClpTxt)
        # image to clipboard
        self.btnClpImg = wx.Button(self.pnlCtrl, -1, label='Image to Clipboard')
        self.Bind(wx.EVT_BUTTON, self.OnClipImage, self.btnClpImg)
        # text to file
        self.btnSavTxt = wx.Button(self.pnlCtrl, -1, label='Output to File')
        self.Bind(wx.EVT_BUTTON, self.OnSaveFile, self.btnSavTxt)

        # window event
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        # set droptarget
        dt = MyFileDropTarget(self.textSrc, self.txtFlname)
        self.textSrc.SetDropTarget(dt)
    
        # font enumerator
        fe = wx.FontEnumerator()
        # we may want only fixed-width ones
        fe.EnumerateFacenames(fixedWidthOnly=True)
        for item in fe.GetFacenames():
            # we don't need vertical fonts 
            if item[0] == '@':
                pass
            else:
                self.choHlFont.Append(item)
        # font size
        for item in ['8','9','10','11','12','14','16','20']:
            self.choFntSiz.Append(item)

        # executable
        if 'nt' in os.name:
            self.hlight = 'c:\\Program Files\\Highlight\\highlight.exe'
        else:
            self.hlight = 'highlight'

        # initialize params
        self.LoadParams()
        # intialize screen scale
        self.InitScale()

        # sizer
        sizer_x = wx.FlexGridSizer(17,2,0,0)
        # file name and scale information
        sizer_x.Add(self.sttFlname, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.txtFlname, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.sttDisply, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.txtDisply, 0, wx.ALL|wx.EXPAND, 4)
        # add space
        sizer_x.Add((20,20), 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add((20,20), 0, wx.ALL|wx.EXPAND, 4)
        # controls
        sizer_x.Add(self.sttSyntax, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.choSyntax, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.sttOutput, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.choOutput, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.sttThemes, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.choThemes, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.sttAstyle, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.choAstyle, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.sttHlFont, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.choHlFont, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.sttFntSiz, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.choFntSiz, 0, wx.ALL|wx.EXPAND, 4)
        # options
        sizer_x.Add(self.sttOption, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.chkLineNo, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add((20,20), 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.chkWrapLn, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add((20,20), 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.chkInLCss, 0, wx.ALL|wx.EXPAND, 4)
        # add space
        sizer_x.Add((20,20), 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add((20,20), 0, wx.ALL|wx.EXPAND, 4)
        # buttons
        sizer_x.Add((20,20), 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.btnConvrt, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add((20,20), 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.btnClpTxt, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add((20,20), 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.btnClpImg, 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add((20,20), 0, wx.ALL|wx.EXPAND, 4)
        sizer_x.Add(self.btnSavTxt, 0, wx.ALL|wx.EXPAND, 4)
        # fit inside the panel
        self.pnlCtrl.SetSizer(sizer_x)

        # sizer for the control panel
        sizer_c = wx.BoxSizer(wx.VERTICAL)
        sizer_c.Add(self.pnlCtrl, 1, wx.ALL|wx.EXPAND, 4)

        # sizer for the frame
        sizer_f = wx.BoxSizer(wx.HORIZONTAL)
        sizer_f.Add(self.spw, 1, wx.ALL|wx.EXPAND, 4)
        sizer_f.Add(sizer_c, 0, wx.ALL|wx.EXPAND, 4)

        self.SetSizer(sizer_f)
        self.SetAutoLayout(1)
        self.Show()

    ## Compile options and call Highlight
    def OnConvert(self, evt):

        # get the selected text region if any
        sel = self.textSrc.GetStringSelection()

        # if none then select entire document instead
        if sel == '':
            sel = self.textSrc.GetValue()

        # no source to convert?
        if sel == '':
            return

        # construct command line with options
        cmd = self.hlight

        # syntax
        try:
            item = self.syntax[self.choSyntax.GetStringSelection()]
        except:
            pass
        else:
            cmd = cmd + ' --syntax=' + item

        # output format
        try:
            item = self.output[self.choOutput.GetStringSelection()]
        except:
            pass
        else:
            cmd = cmd + ' --out-format=' + item

        # astyle
        try:
            item = self.astyle[self.choAstyle.GetStringSelection()]
        except:
            pass
        else:
            if item == ' ':
                pass
            else:
                cmd = cmd + ' --reformat=' + item

        # theme
        try:
            item = self.themes[self.choThemes.GetStringSelection()]
        except:
            pass
        else:
            cmd = cmd + ' --style=' + item

        # font
        facename =  self.choHlFont.GetStringSelection()
        if facename != '':
            cmd = cmd + ' --font="' + facename + '"'

            # size
            size = self.choFntSiz.GetStringSelection()
            if size != '':
                cmd = cmd + ' --font-size=' + size

        # css option
        cmd = cmd + ' --include-style'

        # line numbering
        if self.chkLineNo.GetValue():
            cmd = cmd + ' --line-numbers'

            # get the first line of the selected region
            first_line = sel.splitlines()[0]

            # find the line number of the first selected line
            for index in range(self.textSrc.GetNumberOfLines()):
                if self.textSrc.GetLineText(index) == first_line:
                    # set the starting line number
                    cmd = cmd + ' --line-number-start=' + str(index+1)
                    break

        # wrap
        if self.chkWrapLn.GetValue():
            cmd = cmd + ' --wrap-simple --wrap-no-numbers'

        # inline CSS
        if self.chkInLCss.GetValue():
            cmd = cmd + ' --inline-css'
    
        # run highlight
        p = run(cmd, stdout=PIPE, stderr=PIPE, input=sel, encoding='ascii')

        # error occurred
        if p.stderr != '':
            # display error message
            wx.MessageBox(p.stderr, 'Conversion failed.', wx.ICON_EXCLAMATION)

        else:
            # render html output
            self.webView.SetPage(p.stdout,'')
            # html source
            self.textOut.SetValue(p.stdout)

    ## Copy html source to clipboard
    def OnClipText(self, evt):

        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(wx.TextDataObject(self.textOut.GetValue()))
            wx.TheClipboard.Close()

    ## Copy screen image to clipboard
    def OnClipImage(self, evt):

        if wx.TheClipboard.Open():
            # get webView rectangle in the screen coordinate
            rect = self.webView.GetScreenRect()
            # apply screen scale
            rect.x = rect.x * self.scale
            rect.y = rect.y * self.scale
            rect.width = rect.width * self.scale
            rect.height = rect.height * self.scale
            
            # prepare screen DC
            dcScr = wx.ScreenDC()
            # prepare memory DC
            bmp = wx.Bitmap(rect.width, rect.height)
            memDC = wx.MemoryDC()
            memDC.SelectObject(bmp)
            # render screen area into the memory DC
            memDC.Blit(0,0,rect.width,rect.height,dcScr,rect.x,rect.y)
            # clear memory DC
            memDC.SelectObject(wx.NullBitmap)

            # copy the bitmap into clipboard
            wx.TheClipboard.SetData(wx.BitmapDataObject(bmp))
            wx.TheClipboard.Close()

    ## Load source file
    def OnLoadSource(self, evt):
        with wx.FileDialog(self, 'Open file',
                style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dlg:

            if dlg.ShowModal() == wx.ID_OK:
                fname = dlg.GetPath()
                self.textSrc.LoadFile(fname)
                # show file name
                if 'nt' in os.name:
                    self.txtFlname.SetValue(fname[fname.rfind('\\')+1:])
                else:
                    self.txtFlname.SetValue(fname[fname.rfind('/')+1:])

    ## New source file is loaded
    def OnSourceName(self, evt):
        # file name
        fname = evt.GetString()
        # file extension
        ext = fname[fname.rfind('.')+1:]

        value = ''

        # search filetype dict for the match
        if ext == 'c':
            # extension .c is not found in the dict
            value = 'c'
        elif ext == 'md' or ext == 'MD':
            # extension .md is not found in the dict
            value = 'md'
        elif fname in self.ftmaps['Filenames']:
            # try filename match first
            value = self.ftmaps['Filenames'][fname]
        elif ext in self.ftmaps['Extensions']:
            # then try file extension match
            value = self.ftmaps['Extensions'][ext]

        # none found
        if value == '':
            return

        # let's update the syntax choice box
        for key2, value2 in self.syntax.items():
            if value == value2:
                self.choSyntax.SetStringSelection(key2)
                break

    # Save the html text
    def OnSaveFile(self, evt):
        if self.textOut.GetValue() == '':
            return

        with wx.FileDialog(self, 'Save to file',
                style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as dlg:

            if dlg.ShowModal() == wx.ID_OK:
                self.textOut.SaveFile(dlg.GetPath())

    ## Initialize parameters
    def InitParams(self):

        # config folders and filetype config file
        cmd = self.hlight + ' --print-config'

        # query config settings
        try:
            p = run(cmd, stdout=PIPE, encoding='ascii')
        except:
            wx.MessageBox('Make sure that Highlight is installed properly',
                    'Highlight Execution Error', wx.ICON_EXCLAMATION)
            return

        # collect output
        output = p.stdout.splitlines()

        for idx, line in enumerate(output):
            # read config file search path
            if line == 'Config file search directories:':
                config_path = output[idx+1]
            # read filetype config file location
            elif line == 'Filetype config file:':
                ftcfg_path = output[idx+1]

        try:
            # open filetype config
            ftcfg = open(ftcfg_path, 'r')
            # enumerate other config files
            if 'nt' in os.name:
                themes = glob.glob(config_path + 'themes\\*.theme')
                syntax = glob.glob(config_path + 'langDefs\\*.lang')
                plugin = glob.glob(config_path + 'plugins\\*.lua')
            else:
                themes = glob.glob(config_path + 'themes/*.theme')
                syntax = glob.glob(config_path + 'langDefs/*.lang')
                plugin = glob.glob(config_path + 'plugins/*.lua')
        except:
            wx.MessageBox('Failed to retrieve config files',
                    'Hightlight Config Read Error', wx.ICON_EXCLAMATION)
            return
        
        # filetype mappings
        self.ftmaps = {'Extensions':{},'Filenames':{}}
        saved = ''

        for line in ftcfg:
            # previous line continues
            if saved != '':
                # concatenate two lines
                line = saved.rstrip('\r\n') + line.strip(' ')
                # clear cache
                saved = ''

            # file extensions
            if 'Lang' in line and 'Extensions' in line:
                # unfinished line
                if line.count('{') != line.count('}'):
                    # something is not right
                    if saved != '':
                        # clear saved to prevent error propagation
                        saved = ''
                    # line continues
                    else:
                        # save current line to join next
                        saved = line

                else:
                    keys = line[line.find('Extensions')+12 : line.rfind(',')]
                    value = line[line.find('Lang')+6 : line.find('",')]
                    for key in keys.split(','):
                        key = key.lstrip('" ').rstrip('" }')
                        self.ftmaps['Extensions'][key] = value

            # file names
            elif 'Lang' in line and 'Filenames' in line:
                # unfinished line
                if line.count('{') != line.count('}'):
                    # something is not right
                    if saved != '':
                        # clear saved to prevent error propagation
                        saved = ''
                    # line continues
                    else:
                        # save current line to join next
                        saved = line

                else:
                    keys = line[line.find('Filenames')+11 : line.rfind('"}') + 1]
                    value = line[line.find('Lang')+6 : line.find('",')]
                    for key in keys.split(','):
                        key = key.lstrip('" ').rstrip('" ')
                        self.ftmaps['Filenames'][key] = value

            # ignore shebang
            elif 'Lang' in line and 'Shebang' in line:
                pass

            # ignore empty line
            elif line == '':
                pass

        # themes
        self.themes = {}
        for item in themes:
            # description as key
            self.themes[self.GetDescription(item)] = self.GetFileName(item)

        # syntax (langDefs)
        self.syntax = {}
        for item in syntax:
            # description as key
            self.syntax[self.GetDescription(item)] = self.GetFileName(item)
        
        # plugin
        self.plugin = {}
        for item in plugin:
            # file name as key
            self.plugin[self.GetFileName(item)] = self.GetDescription(item)

        # outputs formats
        self.output = {'html':'html','xhtml':'xhtml','latex':'latex',
                'tex':'tex','odt':'odt','rtf':'rtf','ansi':'ansi','svg':'svg',
                'xterm256':'xterm256','truecolor':'truecolor','pango':'pango',
                'bbcode':'bbcode'}

        # reformat
        self.astyle = {' ':' ','allman':'allman','banner':'banner',
                'gnu':'gnu', 'horstmann':'horstmann','java':'java','kr':'kr',
                'linux':'linux','mozilla':'mozilla','pico':'pico','lisp':'lisp'}

        # option
        self.settng = {'themes':'vim molokai', 'syntax':'C and C++',
                'output':'html', 'astyle':' ', 'plugin':None,
                'hlfont': 'Courier New', 'fntsiz':'10', 
                'option':{'lineno':1, 'wrapln':0, 'inlcss':0} }


    ## Load parameters
    def LoadParams(self):
        try:
            f = open('wxhighlight.cfg', 'rb')
        except:
            wx.MessageBox('Parametes Not found...  Initializing...',
                    'Parameter Load Error', wx.ICON_INFORMATION)
            # if not found then initialize it
            self.InitParams()
        else:
            (self.ftmaps, self.syntax, self.themes, self.plugin, self.output,
                    self.astyle, self.settng) = pickle.load(f)

        # controls should be updated by the parameters
        self.UpdateControls()

    ## Save parameters
    def SaveParams(self):
        # parameters should be updated by the controls
        self.UpdateSettings()

        try:
            f = open('wxhighlight.cfg', 'wb')
        except:
            wx.MessageBox('Failed to save wxhighlight.cfg',
                    'Parameter Save Error', wx.ICON_EXCLAMATION)
        else:
            pickle.dump((self.ftmaps, self.syntax, self.themes, self.plugin,
                self.output, self.astyle, self.settng), f)

    ## Discovery screen scale
    #  \note This does not work on the second monitor
    def InitScale(self):

        # This issue is unique in Windows
        if 'nt' in os.name:
            import multiprocessing as mp
            ctx = mp.get_context('spawn')
            q = ctx.Queue()
            p = ctx.Process(target=SystemMetrics, args=(q,))
            p.start()
            sm = q.get()
            self.scale = sm / ctypes.windll.user32.GetSystemMetrics(0)
            p.join()
        # in Linux it should work without the scaling
        else:
            self.scale = 1.0
        # show the scale
        self.txtDisply.SetLabel('{:.2f}'.format(self.scale))


    ## Fill the choice boxes and set the option check boxes
    def UpdateControls(self):

        for key in self.syntax.keys():
            self.choSyntax.Append(key)
        for key in self.output.keys():
            self.choOutput.Append(key)
        for key in self.themes.keys():
            self.choThemes.Append(key)
        for key in self.astyle.keys():
            self.choAstyle.Append(key)

        if self.settng['syntax']:
            self.choSyntax.SetStringSelection(self.settng['syntax'])
        if self.settng['output']:
            self.choOutput.SetStringSelection(self.settng['output'])
        if self.settng['themes']:
            self.choThemes.SetStringSelection(self.settng['themes'])
        if self.settng['astyle']:
            self.choAstyle.SetStringSelection(self.settng['astyle'])
        if self.settng['hlfont']:
            self.choHlFont.SetStringSelection(self.settng['hlfont'])
        if self.settng['fntsiz']:
            self.choFntSiz.SetStringSelection(self.settng['fntsiz'])

        if self.settng['option']:
            self.chkLineNo.SetValue(self.settng['option']['lineno'])
            self.chkWrapLn.SetValue(self.settng['option']['wrapln'])
            self.chkInLCss.SetValue(self.settng['option']['inlcss'])

        self.Refresh()


    ## Update settings dict
    def UpdateSettings(self):

        self.settng['syntax'] = self.choSyntax.GetStringSelection()
        self.settng['output'] = self.choOutput.GetStringSelection()
        self.settng['themes'] = self.choThemes.GetStringSelection()
        self.settng['astyle'] = self.choAstyle.GetStringSelection()
        self.settng['hlfont'] = self.choHlFont.GetStringSelection()
        self.settng['fntsiz'] = self.choFntSiz.GetStringSelection()
        self.settng['option']['lineno'] = self.chkLineNo.GetValue()
        self.settng['option']['wrapln'] = self.chkWrapLn.GetValue()
        self.settng['option']['inlcss'] = self.chkInLCss.GetValue()


    ## Get the Description string from the file
    def GetDescription(self, fname):
        try:
            f = open(fname, 'r')
        except:
            return ''
        else:
            for line in f:
                if 'Description' in line:
                    return line[line.find('"')+1:line.rfind('"')]


    ## Get the Filename from the path string
    def GetFileName(self, fpath):
        if 'nt' in os.name:
            return fpath[fpath.rfind('\\')+1:fpath.rfind('.')]
        else:
            return fpath[fpath.rfind('/')+1:fpath.rfind('.')]

    ## wx.EVT_CLOSE handler
    def OnClose(self, evt):
        self.SaveParams()
        evt.Skip()


if __name__=='__main__':
    app = wx.App()
    frame = MyFrame(None, -1, "Highlight wxPython GUI", size=(1100,800))
    app.MainLoop()

