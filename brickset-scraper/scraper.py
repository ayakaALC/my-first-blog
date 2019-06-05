import scrapy


class BrickSetSpider(scrapy.Spider):
    name = "brickset_spider"
    start_urls = ['http://vcmsqprd1chd/reports/report/ConfigMgr_VCH/CSSS%20Reports/CSSS%20Custom%20Reports/Check%20Asset%20for%20installed%20applications%20and%20logged%20on%20user%20details']

    def parse(self, response):
        SET_SELECTOR = '.set'
        for brisket in response.css(SET_SELECTOR):
            pass
