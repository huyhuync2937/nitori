using System;
using System.Runtime.InteropServices;
using System.Windows.Input;

namespace SisTCT
{
	public sealed class Win32
	{
		public const uint WS_EX_CONTEXTHELP = 1024u;

		public const uint WS_MINIMIZEBOX = 131072u;

		public const uint WS_MAXIMIZEBOX = 65536u;

		public const int GWL_STYLE = -16;

		public const int GWL_EXSTYLE = -20;

		public const int SWP_NOSIZE = 1;

		public const int SWP_NOMOVE = 2;

		public const int SWP_NOZORDER = 4;

		public const int SWP_FRAMECHANGED = 32;

		public const int WM_SYSCOMMAND = 274;

		public const int SC_CONTEXTHELP = 61824;

		private const byte KEYEVENTF_KEYUP = 2;

		public const int RF_PROCESSMESSAGE = 41251;

		public const int RF_PROCESSWAITINGSHOW = 41254;

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool SetForegroundWindow(IntPtr hWnd);

		[DllImport("user32.dll")]
		public static extern IntPtr SetActiveWindow(IntPtr hWnd);

		[DllImport("user32.dll")]
		public static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);

		[DllImport("user32.dll")]
		public static extern uint GetWindowLong(IntPtr hwnd, int index);

		[DllImport("user32.dll")]
		public static extern int SetWindowLong(IntPtr hwnd, int index, uint newStyle);

		[DllImport("user32.dll")]
		public static extern bool SetWindowPos(IntPtr hwnd, IntPtr hwndInsertAfter, int x, int y, int width, int height, uint flags);

		[DllImport("user32", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
		private static extern void keybd_event(byte bVk, byte bScan, long dwFlags, long dwExtraInfo);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern int SendMessage(IntPtr hwnd, [MarshalAs(UnmanagedType.U4)] int Msg, IntPtr wParam, IntPtr lParam);

		[DllImport("user32.dll")]
		public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);

		[DllImport("shell32.dll", SetLastError = true)]
		public static extern void SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string AppID);

		public static void SendKey(ModifierKeys keyModify, Key key)
		{
			byte bVk = 18;
			switch (keyModify)
			{
				case ModifierKeys.Alt:
					bVk = 18;
					break;
				case ModifierKeys.Control:
					bVk = 17;
					break;
				case ModifierKeys.Shift:
					bVk = 16;
					break;
				case ModifierKeys.Windows:
					bVk = 91;
					break;
			}
			byte bVk2 = (byte)KeyInterop.VirtualKeyFromKey(key);
			if (keyModify != 0)
			{
				keybd_event(bVk, 0, 0L, 0L);
			}
			keybd_event(bVk2, 0, 0L, 0L);
			if (keyModify != 0)
			{
				keybd_event(bVk, 0, 2L, 0L);
			}
			keybd_event(bVk2, 0, 2L, 0L);
		}
	}

}
