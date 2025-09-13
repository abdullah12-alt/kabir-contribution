namespace Client.Services
{
    public static class StreamExtensions
    {
        public static async Task<byte[]> ToBytesAsync(this Stream stream)
        {
            using var memoryStream = new MemoryStream();
            await stream.CopyToAsync(memoryStream);
            return memoryStream.ToArray();
        }
    }
}