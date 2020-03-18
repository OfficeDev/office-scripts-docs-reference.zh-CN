---
title: Office 脚本 API 参考
description: Office 脚本 JavaScript Api 概述
ms.date: 12/12/2019
ms.openlocfilehash: b3e75cbecac13f41c83f019c681b0682571ead41
ms.sourcegitcommit: 2834ee43a14a4b0953d5fb22749c115a42960e20
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2020
ms.locfileid: "42702539"
---
# <a name="office-scripts-api-reference"></a><span data-ttu-id="91cea-103">Office 脚本 API 参考</span><span class="sxs-lookup"><span data-stu-id="91cea-103">Office Scripts API reference</span></span>

<span data-ttu-id="91cea-104">Office 脚本 API 允许您在 web 上的 Excel 中自动执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="91cea-104">The Office Scripts API lets you automate common tasks in Excel on the web.</span></span> <span data-ttu-id="91cea-105">使用此参考文档可了解有关可用于脚本的类、方法和其他类型的详细信息。</span><span class="sxs-lookup"><span data-stu-id="91cea-105">Use this reference documentation to learn more about the classes, methods, and other types available for your scripts.</span></span> <span data-ttu-id="91cea-106">可通过 Office 脚本访问的所有对象都可以在页面左侧的目录中找到。</span><span class="sxs-lookup"><span data-stu-id="91cea-106">All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.</span></span>

## <a name="common-classes"></a><span data-ttu-id="91cea-107">公共类</span><span class="sxs-lookup"><span data-stu-id="91cea-107">Common classes</span></span>

<span data-ttu-id="91cea-108">以下列表细分了 Office 脚本对象模型的基础知识。</span><span class="sxs-lookup"><span data-stu-id="91cea-108">The following list breaks down the basics of the Office Scripts object model.</span></span> <span data-ttu-id="91cea-109">这将显示常见类以及它们之间的关系。</span><span class="sxs-lookup"><span data-stu-id="91cea-109">This shows the common classes and how they relate to one another.</span></span>

- <span data-ttu-id="91cea-110">[工作簿](/javascript/api/office-scripts/excel/excel.workbook)在[WorksheetCollection](/javascript/api/office-scripts/excel/excel.worksheetcollection)中包含一个或多个[工作表](/javascript/api/office-scripts/excel/excel.worksheet)。</span><span class="sxs-lookup"><span data-stu-id="91cea-110">A [Workbook](/javascript/api/office-scripts/excel/excel.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excel/excel.worksheet) in a [WorksheetCollection](/javascript/api/office-scripts/excel/excel.worksheetcollection).</span></span>
- <span data-ttu-id="91cea-111">[工作表](/javascript/api/office-scripts/excel/excel.worksheet)通过[Range](/javascript/api/office-scripts/excel/excel.range)对象提供对单元格的访问。</span><span class="sxs-lookup"><span data-stu-id="91cea-111">A [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excel/excel.range) objects.</span></span>
- <span data-ttu-id="91cea-112">[区域](/javascript/api/office-scripts/excel/excel.range)代表一组连续的单元格。</span><span class="sxs-lookup"><span data-stu-id="91cea-112">A [Range](/javascript/api/office-scripts/excel/excel.range) represents a group of contiguous cells.</span></span>
- <span data-ttu-id="91cea-113">[区域](/javascript/api/office-scripts/excel/excel.range)用于创建和放置[表](/javascript/api/office-scripts/excel/excel.table)、[图表](/javascript/api/office-scripts/excel/excel.chart)、[形状](/javascript/api/office-scripts/excel/excel.shape)以及其他数据可视化或组织对象。</span><span class="sxs-lookup"><span data-stu-id="91cea-113">[Ranges](/javascript/api/office-scripts/excel/excel.range) are used to create and place [Tables](/javascript/api/office-scripts/excel/excel.table), [Charts](/javascript/api/office-scripts/excel/excel.chart), [Shapes](/javascript/api/office-scripts/excel/excel.shape), and other data visualization or organization objects.</span></span>
- <span data-ttu-id="91cea-114">[工作表](/javascript/api/office-scripts/excel/excel.worksheet)包含单个工作表中存在的这些数据对象（如[ChartCollection](/javascript/api/office-scripts/excel/excel.chartcollection)）的集合。</span><span class="sxs-lookup"><span data-stu-id="91cea-114">A [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet) contains collections of those data objects (such as a [ChartCollection](/javascript/api/office-scripts/excel/excel.chartcollection)) that are present in the individual sheet.</span></span>
- <span data-ttu-id="91cea-115">[工作簿](/javascript/api/office-scripts/excel/excel.workbook)。</span><span class="sxs-lookup"><span data-stu-id="91cea-115">[Workbooks](/javascript/api/office-scripts/excel/excel.workbook).</span></span> <span data-ttu-id="91cea-116">包含整个[工作簿](/javascript/api/office-scripts/excel/excel.workbook)的一些数据对象（如[TableCollection](/javascript/api/office-scripts/excel/excel.tablecollection)）的集合。</span><span class="sxs-lookup"><span data-stu-id="91cea-116">contain collections of some of those data objects (such as a [TableCollection](/javascript/api/office-scripts/excel/excel.tablecollection)) for the entire [Workbook](/javascript/api/office-scripts/excel/excel.workbook).</span></span>

<span data-ttu-id="91cea-117">有关 Office 脚本对象模型的详细信息，请访问[web 上 Excel 中 Office 脚本的脚本基础](/office/dev/scripts/develop/scripting-fundamentals)</span><span class="sxs-lookup"><span data-stu-id="91cea-117">For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)</span></span>

## <a name="see-also"></a><span data-ttu-id="91cea-118">另请参阅</span><span class="sxs-lookup"><span data-stu-id="91cea-118">See also</span></span>

- [<span data-ttu-id="91cea-119">关于 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="91cea-119">About Office Scripts</span></span>](/office/dev/scripts/overview/excel)
- [<span data-ttu-id="91cea-120">在 Excel 网页上记录、编辑和创建 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="91cea-120">Record, edit, and create Office Scripts in Excel on the web</span></span>](/office/dev/scripts/tutorials/excel-tutorial)
- [<span data-ttu-id="91cea-121">Web 上的 Excel 中 Office 脚本的脚本基础</span><span class="sxs-lookup"><span data-stu-id="91cea-121">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](/office/dev/scripts/develop/scripting-fundamentals)
