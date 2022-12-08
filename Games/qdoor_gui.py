
import wx

app = wx.App()

frame = wx.Frame(None, title='Quoridoor', size = (1000, 1000))

panel = wx.Panel(frame, wx.ID_ANY)
start = wx.Button(panel, wx.ID_ANY, 'Start', (100, 10))
start.Bind(wx.EVT_BUTTON, print("Pressed"))
start.SetPosition((500, 500))

frame.Show()

app.MainLoop()
