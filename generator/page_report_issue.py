import crinita as cr

html_code = """<h2>How to report any issue with Portable Spreadsheet</h2>
<p>If you find a bug of any type, please let us know. The simplest way is
to go to the GitHub page of Portable Spreadsheet project and raise the issue there.
If you do not know how to do so, find any manual on GitHub issues.</p>
<h2 id="before-you-raise-a-new-issue">Before you raise a new issue</h2>
<p>Please make sure that there is no existing issue of this type or that
this issue is not the part of any other issue.</p>
<h2 id="proposing-suggestions">Proposing suggestions</h2>
<p>If you have a suggestion of any type (leading to improvement of Portable Spreadsheet),
please drop it there as well. There is various type of issues that you
can raise on GitHub.</p>
<h2 id="acknowledgement">Acknowledgement</h2>
<p>We would like to say thank you to anyone who has raised any pertinent
issue already. We really do appreciate this as this is the leading
force for innovations and improvements.</p>
<h2 id="where-to-find-github-page">Where to find GitHub page</h2>
<p>Portable Spreadsheet GitHub project is located at <a href="https://github.com/david-salac/Portable-spreadsheet-generator">https://github.com/david-salac/Portable-spreadsheet-generator</a></p>
"""

ENTITY = cr.Page(
    title="Report issue",
    url_alias='report-issue',
    large_image_path="images/issue_big.jpg",
    content=html_code,
    menu_position=10
)
