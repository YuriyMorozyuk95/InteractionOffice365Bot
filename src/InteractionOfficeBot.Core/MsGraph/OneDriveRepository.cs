using Microsoft.Graph;

namespace InteractionOfficeBot.Core.MsGraph
{
    public class OneDriveRepository
    {
        private readonly GraphServiceClient _graphServiceClient;

        public OneDriveRepository(GraphServiceClient graphServiceClient)
        {
            _graphServiceClient = graphServiceClient;
        }

        public async Task<IDriveItemChildrenCollectionPage> GetRootContents()
        {
            return await _graphServiceClient.Me.Drive.Root.Children.Request().GetAsync(); ;
        }

        public async Task<IDriveItemChildrenCollectionPage> GetFolderContents(string folderPath)
        {
            return await _graphServiceClient.Me.Drive.Root.ItemWithPath(folderPath).Children.Request().GetAsync(); ;
        }

        public async Task RemoveFile(string filePath)
        {
            await _graphServiceClient.Me.Drive.Root.ItemWithPath(filePath).Request().DeleteAsync();
        }

        public async Task<IEnumerable<DriveItem>> SearchOneDrive(string searchText)
        {
            var searchResults = await _graphServiceClient.Me.Drive.Search(searchText).Request().Select("Id").GetAsync();

            var batchRequestContent = new BatchRequestContent();
            List<string> requestIds = new List<string>();

            var pageIterator = PageIterator<DriveItem>.CreatePageIterator(
                _graphServiceClient,
                searchResults,
                driveItem =>
                {
                    var request = _graphServiceClient.Me.Drive.Items[driveItem.Id].Request();
                    var requestId = batchRequestContent.AddBatchRequestStep(request);

                    requestIds.Add(requestId);

                    return true;
                });

            await pageIterator.IterateAsync();

            var batchResponse = await _graphServiceClient.Batch.Request().PostAsync(batchRequestContent);

            var results = new List<DriveItem>();
            foreach (var requestId in requestIds)
            {
                var driveItem = await batchResponse.GetResponseByIdAsync<DriveItem>(requestId);
                results.Add(driveItem);
            }

            return results;
        }

        public async Task<DriveItem> GetFile(string filePath)
        {
            return await _graphServiceClient.Me.Drive.Root.ItemWithPath(filePath).Request().GetAsync();
        }
    }
}
