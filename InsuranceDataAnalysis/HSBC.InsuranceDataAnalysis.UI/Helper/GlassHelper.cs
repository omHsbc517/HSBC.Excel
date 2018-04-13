﻿using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;

[StructLayout(LayoutKind.Sequential)]
public struct MARGINS
{
    public MARGINS(Thickness t)
    {
        Left = (int)t.Left;
        Right = (int)t.Right;
        Top = (int)t.Top;
        Bottom = (int)t.Bottom;
    }
    public int Left;
    public int Right;
    public int Top;
    public int Bottom;
}

public class GlassHelper
{
    [DllImport("dwmapi.dll", PreserveSig = false)]
    static extern void DwmExtendFrameIntoClientArea(
        IntPtr hWnd, ref MARGINS pMarInset);
    [DllImport("dwmapi.dll", PreserveSig = false)]
    static extern bool DwmIsCompositionEnabled();

    public static bool ExtendGlassFrame(Window window, Thickness margin)
    {
        if (!DwmIsCompositionEnabled())
            return false;

        IntPtr hwnd = new WindowInteropHelper(window).Handle;
        if (hwnd == IntPtr.Zero)
            throw new InvalidOperationException(
            "The Window must be shown before extending glass.");

        // Set the background to transparent from both the WPF and Win32 perspectives  
        window.Background = Brushes.Transparent;
        HwndSource.FromHwnd(hwnd).CompositionTarget.BackgroundColor = Colors.Transparent;

        MARGINS margins = new MARGINS(margin);
        DwmExtendFrameIntoClientArea(hwnd, ref margins);
        return true;
    }

    public static bool ExtendGlassFrame(IntPtr hwnd, Thickness margin)
    {
        if (!DwmIsCompositionEnabled())
            return false;

        //IntPtr hwnd = new WindowInteropHelper(window).Handle;
        if (hwnd == IntPtr.Zero)
            throw new InvalidOperationException(
            "The Window must be shown before extending glass.");

        // Set the background to transparent from both the WPF and Win32 perspectives  
        //hwnd.Background = Brushes.Transparent;
        //HwndSource.FromHwnd(hwnd).CompositionTarget.BackgroundColor = Colors.Transparent;

        MARGINS margins = new MARGINS(margin);
        //margins.Top = 0;
        //margins.Right = 1;
        DwmExtendFrameIntoClientArea(hwnd, ref margins);
        return true;
    }

    //private void InitializeFrostedGlass(UIElement glassHost)
    //{
    //    Visual hostVisual = ElementCompositionPreview.GetElementVisual(glassHost);
    //    Compositor compositor = hostVisual.Compositor;
    //    var backdropBrush = compositor.CreateHostBackdropBrush();
    //    var glassVisual = compositor.CreateSpriteVisual();
    //    glassVisual.Brush = backdropBrush;
    //    ElementCompositionPreview.SetElementChildVisual(glassHost, glassVisual);
    //    var bindSizeAnimation = compositor.CreateExpressionAnimation("hostVisual.Size");
    //    bindSizeAnimation.SetReferenceParameter("hostVisual", hostVisual);
    //    glassVisual.StartAnimation("Size", bindSizeAnimation);
    //}

}