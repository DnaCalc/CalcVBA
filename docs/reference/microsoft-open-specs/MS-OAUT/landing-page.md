---
layout: Conceptual
title: '[MS-OAUT]: OLE Automation Protocol | Microsoft Learn'
canonicalUrl: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/bbb05720-f724-45c7-8d17-f83c3d1a3961
ms.service: openspecs-windows
ROBOTS: INDEX, FOLLOW
uhfHeaderId: MSDocsHeader-OpenSpecs
ms.topic: reference
ms.author: cindyle
protocol_rendering: true
description: Specifies the OLE Automation Protocol, which uses DCOM as its transport layer and provides support for an additional set of
locale: en-us
document_id: b1d5e81f-6a77-e608-3413-67a8020085f1
document_version_independent_id: f9123c6c-fd51-93a7-b009-489e140c7cd1
updated_at: 2024-10-30T21:10:00.0000000Z
original_content_git_url: https://github.com/MicrosoftDocs/open_specs_windows/blob/live/documentation/windows_protocols/MS-OAUT/bbb05720-f724-45c7-8d17-f83c3d1a3961.md
gitcommit: https://github.com/MicrosoftDocs/open_specs_windows/blob/817392bbb417d3771ef4fdfbbc11e4226d76470f/documentation/windows_protocols/MS-OAUT/bbb05720-f724-45c7-8d17-f83c3d1a3961.md
git_commit_id: 817392bbb417d3771ef4fdfbbc11e4226d76470f
site_name: Docs
depot_name: MSDN.openspecswindows
page_type: conceptual
toc_rel: toc.json
feedback_system: None
feedback_product_url: ''
feedback_help_link_type: ''
feedback_help_link_url: ''
word_count: 931
asset_id: windows_protocols/ms-oaut/bbb05720-f724-45c7-8d17-f83c3d1a3961
moniker_range_name: 
monikers: []
item_type: Content
source_path: documentation/windows_protocols/MS-OAUT/bbb05720-f724-45c7-8d17-f83c3d1a3961.md
cmProducts:
- https://authoring-docs-microsoft.poolparty.biz/devrel/43093068-2dda-408b-b3fe-dfd705c84f78
- https://authoring-docs-microsoft.poolparty.biz/devrel/bcbcbad5-4208-4783-8035-8481272c98b8
spProducts:
- https://authoring-docs-microsoft.poolparty.biz/devrel/e453d60d-ba7e-43bc-8028-ec38e6b62512
- https://authoring-docs-microsoft.poolparty.biz/devrel/43b2e5aa-8a6d-4de2-a252-692232e5edc8
platformId: 9f7ce283-de7d-f03c-5cb6-d7acc548b591
---

# [MS-OAUT]: OLE Automation Protocol | Microsoft Learn

Specifies the OLE Automation Protocol, which uses DCOM as its transport layer and provides support for an additional set of types as well as for a late-bound calling mechanism.

This page and associated content may be updated frequently. We recommend you subscribe to the [RSS feed](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d.rss) to receive update notifications.

## Published Version

| Date | Protocol Revision | Revision Class | Downloads |
| --- | --- | --- | --- |
| 4/23/2024 | 21.0 | Major | [PDF](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d.pdf) | [DOCX](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-240423.docx) | [Diff](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-240423-diff.pdf) |

[Click here to download a zip file of all PDF files for Windows Protocols.](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/Windows_Protocols.zip)

## Previous Versions

| Date | Protocol Revision | Revision Class | Downloads |
| --- | --- | --- | --- |
| 6/25/2021 | 20.0 | Major | [PDF](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-210625.pdf) | [DOCX](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-210625.docx) | [Diff](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-210625-diff.pdf) |
| 4/7/2021 | 19.0 | Major | [PDF](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-210407.pdf) | [DOCX](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-210407.docx) | [Diff](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-210407-diff.pdf) |
| 9/12/2018 | 18.0 | Major | [PDF](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-180912.pdf) | [DOCX](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-180912.docx) | [Diff](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-180912-diff.pdf) |
| 9/15/2017 | 17.0 | Major | [PDF](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-170915.pdf) | [DOCX](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-170915.docx) | [Diff](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-170915-diff.pdf) |
| 6/1/2017 | 16.0 | None | [PDF](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-170601.pdf) | [DOCX](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-170601.docx) | [Diff](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-170601-diff.pdf) |
| 7/14/2016 | 16.0 | None | [PDF](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-160714.pdf) | [DOCX](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-160714.docx) | [Diff](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-160714-diff.pdf) |
| 10/16/2015 | 16.0 | None | [PDF](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-151016.pdf) | [DOCX](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-151016.docx) |
| 6/30/2015 | 16.0 | Major | [PDF](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-150630.pdf) | [DOCX](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-150630.docx) |
| 5/15/2014 | 15.0 | None | [PDF](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-140515.pdf) | [DOCX](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-140515.doc) |
| 2/13/2014 | 15.0 | None | [PDF](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-140213.pdf) | [DOCX](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-140213.doc) |
| 11/14/2013 | 15.0 | None | [PDF](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-131114.pdf) | [DOCX](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-131114.doc) |
| 8/8/2013 | 15.0 | Major | [PDF](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-130808.pdf) | [DOCX](https://winprotocoldocs-bhdugrdyduf5h2e4.b02.azurefd.net/MS-OAUT/%5bMS-OAUT%5d-130808.doc) |
| 1/31/2013 | 14.0 | None |  |
| 10/25/2012 | 14.0 | None |  |
| 7/12/2012 | 14.0 | None |  |
| 3/30/2012 | 14.0 | None |  |
| 12/16/2011 | 14.0 | Major |  |
| 9/23/2011 | 13.1 | None |  |
| 6/17/2011 | 13.1 | Minor |  |
| 5/6/2011 | 13.0 | None |  |
| 3/25/2011 | 13.0 | None |  |
| 2/11/2011 | 13.0 | None |  |
| 1/7/2011 | 13.0 | Major |  |
| 11/19/2010 | 12.0.2 | None |  |
| 10/8/2010 | 12.0.2 | None |  |
| 8/27/2010 | 12.0.2 | None |  |
| 7/16/2010 | 12.0.2 | None |  |
| 6/4/2010 | 12.0.2 | Editorial |  |
| 4/23/2010 | 12.0.1 | Editorial |  |
| 3/12/2010 | 12.0 | Major |  |
| 1/29/2010 | 11.0.1 | Editorial |  |
| 12/18/2009 | 11.0 | Major |  |
| 11/6/2009 | 10.0 | Major |  |
| 9/25/2009 | 9.1 | Minor |  |
| 8/14/2009 | 9.0.1 | Editorial |  |
| 7/2/2009 | 9.0 | Major |  |
| 5/22/2009 | 8.4.2 | Editorial |  |
| 4/10/2009 | 8.4.1 | Editorial |  |
| 2/27/2009 | 8.4 | Minor |  |
| 1/16/2009 | 8.3.1 | Editorial |  |
| 12/5/2008 | 8.3 | Minor |  |
| 10/24/2008 | 8.2 | Minor |  |
| 8/29/2008 | 8.1.1 | Editorial |  |
| 7/25/2008 | 8.1 | Minor |  |
| 6/20/2008 | 8.0 | Major |  |
| 5/16/2008 | 7.0 | Major |  |
| 3/14/2008 | 6.0 | Major |  |
| 1/25/2008 | 5.0 | Major |  |
| 11/30/2007 | 4.0 | Major |  |
| 10/23/2007 | 3.0 | Major |  |
| 9/28/2007 | 2.1 | Minor |  |
| 8/10/2007 | 2.0.1 | Editorial |  |
| 7/20/2007 | 2.0 | Major |  |
| 6/1/2007 | 1.2.1 | Editorial |  |
| 5/11/2007 | 1.2 | Minor |  |
| 4/3/2007 | 1.1 | Minor |  |
| 3/2/2007 | 1.0 | Major |  |
| 12/18/2006 | 0.1 | New |  |

## Preview Versions

From time to time, Microsoft may publish a preview, or pre-release, version of an Open Specifications technical document for community review and feedback. To submit feedback for a preview version of a technical document, please follow any instructions specified for that document. If no instructions are indicated for the document, please provide feedback by using the [Open Specification Forums](https://aka.ms/AA9oo6c).

The preview period for a technical document varies. Additionally, not every technical document will be published for preview.

A preview version of this document may be available on the [Windows Protocols - Preview Documents](../ms-winprotlp/8a9c667b-2825-46a8-8066-a80681233c33) page. After the preview period, the most current version of the document is available on this page.

## Development Resources

Find resources for creating interoperable solutions for Microsoft software, services, hardware, and non-Microsoft products: 

[Plugfests and Events](https://msdn.microsoft.com/en-us/openspecifications/dn750988), [Test Tools](https://msdn.microsoft.com/en-us/openspecifications/dn750986), [Development Support](https://msdn.microsoft.com/en-us/openspecifications/cc816063), and [Open Specifications Dev Center](https://msdn.microsoft.com/en-us/openspecifications).

## Intellectual Property Rights Notice for Open Specifications Documentation

- **Technical Documentation. **Microsoft publishes Open Specifications documentation (“this documentation”) for protocols, file formats, data portability, computer languages, and standards support. Additionally, overview documents cover inter-protocol relationships and interactions.
- **Copyrights**. This documentation is covered by Microsoft copyrights. Regardless of any other terms that are contained in the terms of use for the Microsoft website that hosts this documentation, you can make copies of it in order to develop implementations of the technologies that are described in this documentation and can distribute portions of it in your implementations that use these technologies or in your documentation as necessary to properly document the implementation. You can also distribute in your implementation, with or without modification, any schemas, IDLs, or code samples that are included in the documentation. This permission also applies to any documents that are referenced in the Open Specifications documentation.
- **No Trade Secrets**. Microsoft does not claim any trade secret rights in this documentation.
- **Patents**. Microsoft has patents that might cover your implementations of the technologies described in the Open Specifications documentation. Neither this notice nor Microsoft's delivery of this documentation grants any licenses under those patents or any other Microsoft patents. However, a given Open Specifications document might be covered by the Microsoft [Open Specifications Promise](https://go.microsoft.com/fwlink/?LinkId=214445) or the [Microsoft Community Promise](https://go.microsoft.com/fwlink/?LinkId=214448). If you would prefer a written license, or if the technologies described in this documentation are not covered by the Open Specifications Promise or Community Promise, as applicable, patent licenses are available by contacting iplg@microsoft.com.
- **License Programs**. To see all of the protocols in scope under a specific license program and the associated patents, visit the [Patent Map](https://aka.ms/AA9ufj8).
- **Trademarks**. The names of companies and products contained in this documentation might be covered by trademarks or similar intellectual property rights. This notice does not grant any licenses under those rights. For a list of Microsoft trademarks, visit https://www.microsoft.com/trademarks.
- **Fictitious Names**. The example companies, organizations, products, domain names, email addresses, logos, people, places, and events that are depicted in this documentation are fictitious. No association with any real company, organization, product, domain name, email address, logo, person, place, or event is intended or should be inferred.

**Reservation of Rights**. All other rights are reserved, and this notice does not grant any rights other than as specifically described above, whether by implication, estoppel, or otherwise. 

**Tools**. The Open Specifications documentation does not require the use of Microsoft programming tools or programming environments in order for you to develop an implementation. If you have access to Microsoft programming tools and environments, you are free to take advantage of them. Certain Open Specifications documents are intended for use in conjunction with publicly available standards specifications and network programming art and, as such, assume that the reader either is familiar with the aforementioned material or has immediate access to it.

**Support.** For questions and support, please contact dochelp@microsoft.com.