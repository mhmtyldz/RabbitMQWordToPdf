using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using RabbitMQ.Client;
using WordToPdf.Producer.Models;

namespace WordToPdf.Producer.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public IActionResult WordToPdfPage()
        {
            return View();
        }
        [HttpPost]
        public IActionResult WordToPdfPage(WordToPdfModel wordToPdf)
        {
            var rabbitMQConnectionString = _configuration.GetConnectionString("RabbitMQLocalString");
            var factory = new ConnectionFactory() { HostName = rabbitMQConnectionString };

            using (var connection = factory.CreateConnection())
            {
                using (var channel = connection.CreateModel())
                {
                    //Direct exchangeci hazırladık.
                    channel.ExchangeDeclare(exchange: "convert-exchange", type: ExchangeType.Direct, durable: true, autoDelete: false, arguments: null);

                    //Kuyruğu normalde consumer da oluşturuyodum şuan burda oluşturuyorum sebebi ise :
                    //Bu oluşturduğumuz mesajları dinleyen bir subscriber olmasa dahi kuyruğumun oluşmasını istiyorum

                    //exclusive birden fazla bağlantı bu kuyruğu kullansın diye
                    channel.QueueDeclare(queue: "File", durable: true, exclusive: false, autoDelete: false, arguments: null);

                    //Kuyruğu burada bind ediyorum mesajımın kaybolmasını istemiyorum.
                    channel.QueueBind(queue: "File", exchange: "convert-exchange", routingKey: "WordToPdf");
                    var messageWordToPdf = new MessageWordToPdf();
                    using (MemoryStream ms = new MemoryStream())
                    {
                        wordToPdf.WordFile.CopyTo(ms);
                        messageWordToPdf.WordByte = ms.ToArray();
                        messageWordToPdf.Email = wordToPdf.Email;
                        messageWordToPdf.FileName = Path.GetFileName(wordToPdf.WordFile.FileName);

                        string serializeMessage = JsonConvert.SerializeObject(messageWordToPdf);

                        var byteMessage = Encoding.UTF8.GetBytes(serializeMessage);




                        //Mesajımı sağlama almak için aşağıdaki işlemleri yapıyorum kuyruğumu yukarda durable true dedik.

                        var properties = channel.CreateBasicProperties();

                        properties.Persistent = true; //Mesajlarım Rabbitmq restart atsa dahi kayboblmayacak

                        channel.BasicPublish(exchange: "convert-exchange", routingKey: "WordToPdf", basicProperties: properties, byteMessage);

                        ViewBag.Result = "Word dosyanız pdf dosyasına dönüştürüşdükten sonra size email olarak gönderilecektir";

                        return View();

                    }

                }
            }



        }
    }
}
