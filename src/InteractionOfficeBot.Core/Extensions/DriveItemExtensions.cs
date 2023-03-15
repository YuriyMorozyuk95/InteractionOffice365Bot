using Microsoft.Graph;

namespace InteractionOfficeBot.Core.Extensions
{
    public static class DriveItemExtensions
    {
        public static string GetFullPath(this DriveItem driveItem)
        {
            var folderPath = driveItem.ParentReference?.Path.Substring(driveItem.ParentReference.Path.IndexOf(':') + 1);
            var fileName = driveItem.GetDisplayName();

            return string.Join("\\", new[] { folderPath, fileName }.Where(x => !string.IsNullOrEmpty(x)));
        }

        public static string GetDisplayName(this DriveItem driveItem)
        {
            return driveItem.Folder != null ? driveItem.Name : $"{driveItem.Name}\\";
        }

        public static bool IsFileOrFolder(this DriveItem driveItem)
        {
            return driveItem.File != null || driveItem.Folder != null;
        }
    }
}
