---
layout: post
title:  "Index and Match"
date:   2021-06-24 23:42:05 +0530
categories: 
---

The [INDEX](https://support.google.com/docs/answer/3098242?hl=en) and [MATCH](https://support.google.com/docs/answer/3093378?hl=en) functions combined is often a faster and more stable alternative to the popular [VLOOKUP](https://support.google.com/docs/answer/3093318?hl=en) on Sheets. A very basic implementation is outlined in the below mentioned formula

{% highlight ruby %}
=INDEX(IMPORTRANGE(properties!$C$4,properties!$C$5&"!D2:N"), MATCH(A13,IMPORTRANGE(properties!$C$4,properties!$C$5&"!B2:B"), 0))
{% endhighlight %}

First thing to note would be that we are not plugging in the sheet ID and the tab name directly into the formula. Instead this can be stored in a separate tab on the parent sheet. Over here the formula pulls that information from a tab called `properties`. 

> This separation helps in making the formula more manageable and also makes implementing changes much easier.

The formula will check if the value in cell `A13` matches the `B` column of any particular row. If it does then it retrieves columns `D to N` of that row.

In case there is no match found then a <mark>#NA</mark> error is reported, this can be handled using the [ISNA](https://support.google.com/docs/answer/9365944?hl=en) function