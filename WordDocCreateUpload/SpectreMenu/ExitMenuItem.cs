namespace WordDocCreateUpload.SpectreMenu
{
    class ExitMenuItem : MenuItem
    {
        public ExitMenuItem()
        {
            setName("[red]Exit[/]");
        }
        public override Task<IMenuItem> navigate()
        {
            System.Environment.Exit(0);
            return null;
        }
    }
}
