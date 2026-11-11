using System;
using System.IO;
using System.Runtime.InteropServices;

internal static class Ornek
{
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    private delegate int Satir(IntPtr basliklar, IntPtr hucreler, int sutunSayisi,
        IntPtr secim, int secimSayisi, int baslikKipi, IntPtr cikti, int kapasite);

    private static string Oku(Satir satir, string[] basliklar, string[] hucreler)
    {
        int adet = hucreler.Length;
        var ayrilanlar = new System.Collections.Generic.List<IntPtr>();
        IntPtr Ayir(int bayt)
        {
            IntPtr adres = Marshal.AllocHGlobal(bayt);
            ayrilanlar.Add(adres);
            return adres;
        }
        IntPtr Dizi(string[] metinler)
        {
            IntPtr dizi = Ayir(IntPtr.Size * adet);
            for (int i = 0; i < adet; i++)
            {
                IntPtr metin = Marshal.StringToHGlobalUni(metinler[i]);
                ayrilanlar.Add(metin);
                Marshal.WriteIntPtr(dizi, i * IntPtr.Size, metin);
            }
            return dizi;
        }
        try
        {
            IntPtr b = Dizi(basliklar);
            IntPtr h = Dizi(hucreler);
            int boyut = satir(b, h, adet, IntPtr.Zero, 0, 2, IntPtr.Zero, 0);
            if (boyut <= 0) throw new InvalidOperationException($"DLL sonucu {boyut}");
            IntPtr cikti = Ayir(boyut * 2);
            if (satir(b, h, adet, IntPtr.Zero, 0, 2, cikti, boyut) != boyut)
                throw new InvalidOperationException("Satır hazırlanamadı.");
            return Marshal.PtrToStringUni(cikti, boyut - 1) ?? "";
        }
        finally
        {
            foreach (IntPtr adres in ayrilanlar) Marshal.FreeHGlobal(adres);
        }
    }

    private static void Main(string[] args)
    {
        if (RuntimeInformation.ProcessArchitecture != Architecture.X64 &&
            RuntimeInformation.ProcessArchitecture != Architecture.X86)
            throw new PlatformNotSupportedException("x64 veya x86 .NET kullanın.");
        string mimari = Environment.Is64BitProcess ? "x64" : "x86";
        string dizin = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();
        IntPtr dll = NativeLibrary.Load(Path.GetFullPath(Path.Combine(dizin, mimari, "TabloOkuma.dll")));
        try
        {
            var satir = Marshal.GetDelegateForFunctionPointer<Satir>(
                NativeLibrary.GetExport(dll, "Tablo_Satir"));
            Console.WriteLine(Oku(satir, new[] { "İsim", "Numara", "Sıra" },
                new[] { "Emrah", "2121", "23" }));
        }
        finally { NativeLibrary.Free(dll); }
    }
}
