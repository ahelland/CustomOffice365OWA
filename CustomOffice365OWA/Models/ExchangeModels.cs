using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace CustomOffice365OWA.Models
{
    #region Email
    public class Message_odataContext
    {
        [JsonProperty("@odata.context")]
        public string context;
        [JsonProperty("value")]
        public List<Message> Messages { get; set; }
    }

    public class Message
    {
        [JsonProperty("@odata.id")]
        [ScaffoldColumn(false)]
        public string odataId { get; set; }
        [JsonProperty("@odata.editLink")]
        [ScaffoldColumn(false)]
        public string odataEditLink { get; set; }
        [ScaffoldColumn(false)]
        public string Id { get; set; }
        [ScaffoldColumn(false)]
        public string ChangeKey { get; set; }
        [ScaffoldColumn(false)]
        public string ConversationId { get; set; }
        [ScaffoldColumn(false)]
        public string ParentFolderId { get; set; }
        [JsonProperty("Attachments@odata.navigationLink")]
        [ScaffoldColumn(false)]
        public string AttachmentLink { get; set; }
        [ScaffoldColumn(false)]
        public string EventId { get; set; }

        public DateTime DateTimeReceived { get; set; }
        public DateTime DateTimeSent { get; set; }
        public DateTime DateTimeCreated { get; set; }
        public DateTime LastModifiedTime { get; set; }
        public Recipient From { get; set; }
        public Recipient Sender { get; set; }
        public Recipient[] ToRecipients { get; set; }
        public string Subject { get; set; }
        public string BodyPreview { get; set; }
        public Body Body { get; set; }
        public string Importance { get; set; }

        public Recipient[] CcRecipients { get; set; }
        public Recipient[] BccRecipients { get; set; }
        public Recipient[] ReplyTo { get; set; }
        public string IsDeliveryReceiptRequested { get; set; }
        public string IsReadReceiptRequested { get; set; }
        public string IsDraft { get; set; }
        public string IsRead { get; set; }
        public string MeetingMessageType { get; set; }
        public string HasAttachments { get; set; }
        public Category[] Categories { get; set; }
        public string ClassName { get; set; }
    }

    public class Recipient
    {
        public string Name { get; set; }
        public string Address { get; set; }
    }

    public class Body
    {
        public string ContentType { get; set;}
        public string Content {get;set;}
    }

    public class Category
    {
        public string Name {get;set;}
    }

#endregion
}