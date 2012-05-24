DocViewer
=========

view office document online,without office installed.also support pdf,images and plain text files
=========

                  Word    save as           swftools
office documents  PPT     ------->    pdf   -------->    swf
                  Excel
=========
tutorial
=========
1.office 2007 with SaveAsPdf plugin must be installed on the file server.(office 2010 is recommended)
2.download the latest Quartz.Net from https://github.com/quartznet/quartznet,and install it as Windows Service.
3.download the latest swftools from www.swftools.org,install it.
4.add ToPDFJob ToSWFJob to the quartz_jobs.xml