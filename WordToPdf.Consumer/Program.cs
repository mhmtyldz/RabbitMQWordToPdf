using Newtonsoft.Json;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using Spire.Doc;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace WordToPdf.Consumer
{
    class Program
    {
        static void Main(string[] args)
        {
            bool result = false;
            var factory = new ConnectionFactory() { HostName = "localhost" };

            using (var connection = factory.CreateConnection())
            {
                using (var channel = connection.CreateModel())
                {
                    //Direct exchangeci hazırladık.
                    channel.ExchangeDeclare(exchange: "convert-exchange", type: ExchangeType.Direct, durable: true, autoDelete: false, arguments: null);

                    channel.QueueBind(queue: "File", exchange: "convert-exchange", "WordToPdf");

                    channel.BasicQos(0, 1, false);

                    var consumer = new EventingBasicConsumer(channel);

                    channel.BasicConsume("File", false, consumer);

                    consumer.Received += (model, ea) =>
                    {
                        try
                        {
                            Console.WriteLine("Kuyruktan bir mesaj alındı ve işleniyor");

                            Document document = new Document();

                            string message = Encoding.UTF8.GetString(ea.Body.Span);

                            MessageWordToPdf messageWordToPdf = JsonConvert.DeserializeObject<MessageWordToPdf>(message);

                            document.LoadFromStream(new MemoryStream(messageWordToPdf.WordByte), FileFormat.Docx2013);

                            using (MemoryStream ms = new MemoryStream())
                            {
                                document.SaveToStream(ms, FileFormat.PDF);

                                result = EmailSend(messageWordToPdf.Email, ms, messageWordToPdf.FileName);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Hata meydana geldi:" + ex.Message);
                        }

                        if (result)
                        {
                            Console.WriteLine("Kuyruktan Mesaj başarıla işlendi...");
                            channel.BasicAck(ea.DeliveryTag, false);
                        }
                    };
                }
            }








        }

        public static bool EmailSend(string email, MemoryStream memoryStream, string fileName)
        {
            try
            {
                memoryStream.Position = 0;

                System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Application.Pdf);

                Attachment attachment = new Attachment(memoryStream, ct);

                attachment.ContentDisposition.FileName = $"{fileName}.pdf";

                MailMessage mailMessage = new MailMessage();

                //SmtpClient smtpClient = new SmtpClient();

                mailMessage.From = new MailAddress("yldzmahmut0@gmail.com");

                mailMessage.To.Add(email);

                mailMessage.Subject = "Sitemize Hoşgeldiniz";
                mailMessage.Body = "Pdf dosyanız mail adresine ulaşmıştır";

                mailMessage.IsBodyHtml = true;

                mailMessage.Attachments.Add(attachment);

                using (var client = new SmtpClient())
                {
                    client.Credentials = new NetworkCredential("azure_1d8324f892f39c4a4889c88739ea92a2@azure.com", "G@@gle1907");
                    client.Port = 587;
                    client.Host = "smtp.sendgrid.net";
                    client.EnableSsl = true;
                    client.Send(mailMessage);

                    Console.WriteLine($"Sonuc : Pdf {email} adresine gönderilmiştir");
                    memoryStream.Close();
                    memoryStream.Dispose();
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Mail gönderim sırasında bir hata oluştu. {ex.InnerException}");
                return false;
            }

        }
    }
}
