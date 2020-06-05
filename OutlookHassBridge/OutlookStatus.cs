namespace OutlookHassBridge
{
    public class OutlookStatus
    {
        public bool? OutlookUnread { get; set; }

        public bool Equals(OutlookStatus obj)
        {
            return obj != null && (obj.OutlookUnread == OutlookUnread);
        }

        public OutlookStatus(bool? outlookUnread)
        {
            OutlookUnread = outlookUnread;
        }
    }
}
