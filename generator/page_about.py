import crinita as cr

# Converted using: https://markdowntohtml.com/

html_code = """<p>Author: David Salac <a href="http://github.com/david-salac">http://github.com/david-salac</a></p>
<p>Simple spreadsheet that keeps tracks of each operation in defined programming
languages. Logic allows export sheets to Excel files (and see how each cell is
computed), to the JSON strings with description of computation e. g. in native
language. Other formats like HTML, CSV and Markdown (MD) are also supported. It
also allows to reconstruct behaviours in native Python with NumPy.</p>
<h2 id="key-components-of-the-library">Key components of the library</h2>
<p>There are five main objects in the library:</p>
<ol>
<li><strong><em>Grammar</em></strong>: the set of rule that defines language.</li>
<li><strong><em>Cell</em></strong>: single cell inside the spreadsheet.</li>
<li><strong><em>Word</em></strong>: the word describing how the cell is created in each language.</li>
<li><strong><em>Spreadsheet</em></strong>: a set of cells with a defined shape.</li>
<li><strong><em>Cell slice</em></strong>: a subset of the spreadsheet. </li>
</ol>
<h3 id="grammar">Grammar</h3>
<p>The grammar defines a context-free language (by Chomsky hierarchy). It is
used for describing each operation that is done with the cell. The typical
world is constructed using prefix, suffix and actual value by creating a
string like &quot;PrefixValueSuffix&quot;. Each supported operation is defined in
grammar (that tells how the word is created when the operation is called).</p>
<p>There are two system languages (grammars): Python and Excel. There is also
one language &quot;native&quot; that describes operations in a native tongue logic.</p>
<h4 id="adding-the-new-grammar">Adding the new grammar</h4>
<p>Operations with grammars are encapsulated in the class <code>GrammarUtils</code>.</p>
<p>Grammar has to be defined as is described in the file <code>grammars.py</code> in the
global variable <code>GRAMMAR_PATTERN</code>. It is basically the dictionary matching
the description.</p>
<p>To validate the grammar (in the variable <code>grammar</code>) use: </p>
<pre><code class="lang-python">is_valid: <span class="hljs-keyword">bool</span> = GrammarUtils.validate_grammar(grammar)
</code></pre>
<p>To add the grammar describing some language (in the variable <code>language</code>)
to the system (in the variable <code>grammar</code>) use: </p>
<pre><code class="lang-python"><span class="hljs-selector-tag">GrammarUtils</span><span class="hljs-selector-class">.add_grammar</span>(<span class="hljs-selector-tag">grammar</span>, <span class="hljs-selector-tag">language</span>)
</code></pre>
<p>User can also check what languages are currently available using the
static method <code>get_languages</code>:</p>
<pre><code class="lang-python"><span class="hljs-symbol">languages_in_the_system:</span> <span class="hljs-keyword">Set</span>[str] = GrammarUtils.get_languages()
</code></pre>
<h3 id="cells">Cells</h3>
<p>It represents the smallest element in the spreadsheet. Cell encapsulates basic
arithmetic and logical operations that are needed. A cell is represented by
the class of the same name <code>Cell</code>. It is highly recommended not to use
this class directly but only through the spreadsheet instance.</p>
<p>Currently, the supported operations are described in the subsections
<em>Computations</em> bellow in this document (as all that unary, binary and other
functions).</p>
<p>The main purpose of the cell is to keep the value (the numerical result of
the computation) and the word (how is an operation or constant represented
in all languages).</p>
<h3 id="words">Words</h3>
<p>Word represents the current computation or value of the cell using in given
languages.</p>
<p>A typical example of the word can be (in language excel):</p>
<pre><code class="lang-python">B2*(<span class="hljs-name">C1+C2</span>)
</code></pre>
<p>The equivalent word in the language Python:</p>
<pre><code class="lang-python">values[<span class="hljs-number">1</span>,<span class="hljs-number">1</span>]*(values[<span class="hljs-number">0</span>,<span class="hljs-number">2</span>]+values[<span class="hljs-number">1</span>,<span class="hljs-number">2</span>])
</code></pre>
<p>Words are constructed using prefixes and suffixes defined by the grammar.
Each language also has some special features that are also described in
the grammar (like whether the first index represents column or a row).</p>
<p>Words are important later when the output is exported to some file in given
format or to JSON.
Operations with words (and word as a data structure) are located in the
class <code>WordConstructor</code>. It should not be used directly.</p>
<h3 id="spreadsheet-class">Spreadsheet class</h3>
<p>The Spreadsheet is the most important class of the whole package. It is
located in the file <code>spreadsheet.py</code>. It encapsulates the functionality
related to accessing cells and modifying them as well as the functionality
for exporting of the computed results to various formats.</p>
<p>Class is strongly motivated by the API of the Pandas DataFrame. </p>
<p>Spreadsheet functionality is documented bellow in a special section.</p>
<h3 id="cell-slice">Cell slice</h3>
<p>Represents the special object that is created when some part slice of the
spreadsheet is created. Basically, it encapsulates the set of cells and
aggregating operations (sum, product, minimum, maximum, average). For example:</p>
<pre><code class="lang-python"><span class="hljs-attr">some_slice</span> = spreadsheet_instance.iloc[<span class="hljs-number">1</span>,:]
<span class="hljs-attr">average_of_slice</span> = some_slice.mean()
</code></pre>
<p>selected the second row in the spreadsheet and compute the average (mean)
of values in the slice.</p>
<p>Cell slice is represented in the class <code>CellSlice</code> in the
file <code>cell_slice.py</code>.</p>
<p>If you want to assign some value to a <code>CellSlice</code> object, you can use
overloaded operator <code>&lt;&lt;=</code></p>
<pre><code class="lang-python">some_slice = spreadsheet_instance.iloc[1,:]
average_of_slice &lt;&lt;= 55.6  # Some assigned value
</code></pre>
<p>However, it is strongly recommended to use standard assigning through
the Spreadsheet object described below.</p>
<h4 id="functionality-of-the-cellslice-class">Functionality of the CellSlice class</h4>
<p>Cell slice is mainly related to the aggregating functions described in
the subsection <em>Aggregate functions</em> bellow.</p>
<p>There is also a functionality related to setting the values to some
constant or reference to another cell. This functionality should not
be used directly.</p>
<p>Cell slices can be exported in the same way as a whole spreadsheet (methods
are discussed below).</p>
<h2 id="spreadsheets-functionality">Spreadsheets functionality</h2>
<p>All following examples expect that user has already imported package.</p>
<pre><code class="lang-python"><span class="hljs-keyword">import</span> portable_spreadsheet <span class="hljs-keyword">as</span> ps
</code></pre>
<p>The default (or system) languages are Excel and Python. There is also
a language called &#39;native&#39; ready to be used.</p>
<h3 id="how-to-create-a-spreadsheet">How to create a spreadsheet</h3>
<p>The easiest function is to use the built-in static method <code>create_new_sheet</code>:</p>
<pre><code class="lang-python">sheet = ps<span class="hljs-selector-class">.Spreadsheet</span><span class="hljs-selector-class">.create_new_sheet</span>(
    number_of_rows, number_of_columns, [rows_columns]
)
</code></pre>
<p>if you wish to include some user-defined languages or the language
called &#39;native&#39; (which is already in the system), you also need to
pass the argument <code>rows_columns</code> (that is a dictionary with keys as
languages and values as lists with column names in a given non-system
language).</p>
<p>For example, if you choose to add <em>&#39;native&#39;</em> language (already available in
grammars), you can use a shorter version:</p>
<pre><code class="lang-python">sheet = ps.Spreadsheet.create<span class="hljs-number">_n</span>ew<span class="hljs-number">_</span>sheet(
    number<span class="hljs-number">_</span><span class="hljs-keyword">of</span><span class="hljs-number">_</span>rows, number<span class="hljs-number">_</span><span class="hljs-keyword">of</span><span class="hljs-number">_</span>columns, 
    {
        <span class="hljs-string">"native"</span>: cell<span class="hljs-number">_</span>indices<span class="hljs-number">_</span>generators[<span class="hljs-string">'native'</span>](number<span class="hljs-number">_</span><span class="hljs-keyword">of</span><span class="hljs-number">_</span>rows, 
                                                    number<span class="hljs-number">_</span><span class="hljs-keyword">of</span><span class="hljs-number">_</span>columns),
    }
)
</code></pre>
<p>Other (keywords) arguments:</p>
<ol>
<li><code>rows_labels (List[Union[str, SkippedLabel]])</code>: <em>(optional)</em> List of masks
(aliases) for row names.</li>
<li><code>columns_labels (List[Union[str, SkippedLabel]])</code>: <em>(optional)</em> List of
masks (aliases) for column names. If the instance of SkippedLabel is
used, the export skips this label.</li>
<li><code>rows_help_text (List[str])</code>: <em>(optional)</em> List of help texts for each row.</li>
<li><code>columns_help_text (List[str])</code>: <em>(optional)</em> List of help texts for each
column. If the instance of SkippedLabel is used, the export skips this label.</li>
<li><code>excel_append_row_labels (bool)</code>: <em>(optional)</em> If True, one column is added
on the beginning of the sheet as a offset for labels.</li>
<li><code>excel_append_column_labels (bool)</code>: <em>(optional)</em> If True, one row is
added on the beginning of the sheet as a offset for labels.</li>
<li><code>warning_logger (Callable[[str], None]])</code>: Function that logs the warnings
(or <code>None</code> if logging should be skipped).</li>
</ol>
<p>First two are the most important because they define labels for the columns
and rows indices. The warnings mention above occurs when the slices are
exported (which can lead to data losses).</p>
<h4 id="how-to-change-the-size-of-the-spreadsheet">How to change the size of the spreadsheet</h4>
<p>You can only expand the size of the spreadsheet (it&#39;s because of the
built-in behaviour of language construction). We, however, strongly recommend
not to do so. Simplified logic looks like:</p>
<pre><code class="lang-python"># Append <span class="hljs-number">7</span> rows <span class="hljs-keyword">and</span> <span class="hljs-number">8</span> <span class="hljs-built_in">columns</span> to existing sheet:
sheet.<span class="hljs-built_in">expand</span>(
    <span class="hljs-number">7</span>, <span class="hljs-number">8</span>,  
    {
        <span class="hljs-string">"native"</span>: ([...], [...])  # Fill <span class="hljs-number">8</span> <span class="hljs-built_in">new</span> <span class="hljs-built_in">values</span> <span class="hljs-keyword">for</span> rows, <span class="hljs-built_in">columns</span> here
    }
)
</code></pre>
<p>Parameters of the <code>Spreadsheet.expand</code> method are of the same
logic and order as the parameters of <code>Spreadsheet.create_new_sheet</code>.</p>
<h3 id="column-and-row-labels">Column and row labels</h3>
<p>Labels are set once when a sheet is created (or expanded in size). If you
want to read them as a tuple of labels, you can use the following properties:</p>
<ul>
<li><code>columns</code>: property that returns labels of columns as a tuple of strings.
It can be called on both slices or directly on <code>Spreadsheet</code> class instances.</li>
<li><code>index</code>: property that returns the labels of rows as a tuple of strings.
It can be called on both slices or directly on <code>Spreadsheet</code> class instances.</li>
</ul>
<p>Example:</p>
<pre><code class="lang-python"><span class="hljs-symbol">column_labels:</span> Tuple[str] = <span class="hljs-keyword">sheet.columns </span> <span class="hljs-comment"># Get the column labels</span>
<span class="hljs-symbol">row_labels:</span> Tuple[str] = <span class="hljs-keyword">sheet.index </span> <span class="hljs-comment"># Get the row labels</span>
</code></pre>
<h3 id="shape-of-the-spreadsheet-object">Shape of the Spreadsheet object</h3>
<p>If you want to know what is the actual size of the spreadsheet, you can
use the property <code>shape</code> that behaves as in Pandas. It returns you the
tuple with a number of rows and number of columns (on the second position).</p>
<h3 id="accessing-setting-the-cells-in-the-spreadsheet">Accessing/setting the cells in the spreadsheet</h3>
<p>You to access the value in the position you can use either the integer
position (indexed from 0) or the label of the row/column.</p>
<pre><code class="lang-python"># Returns the <span class="hljs-built_in">value</span> at <span class="hljs-built_in">second</span> <span class="hljs-built_in">row</span> <span class="hljs-built_in">and</span> third colu<span class="hljs-symbol">mn:</span>
<span class="hljs-built_in">value</span> = sheet.iloc[<span class="hljs-number">1</span>,<span class="hljs-number">2</span>]
# Returns the <span class="hljs-built_in">value</span> by the name of the <span class="hljs-built_in">row</span>, <span class="hljs-built_in">column</span>
<span class="hljs-built_in">value</span> = sheet.loc['super the label of <span class="hljs-built_in">row</span>', '<span class="hljs-built_in">even</span> better label of <span class="hljs-built_in">column</span>']
</code></pre>
<p>As you can see, there are build-in properties <code>loc</code> and <code>iloc</code> for accessing
the values (the <code>loc</code> access based on the label, and <code>iloc</code> access the cell
based on the integer position).</p>
<p>The same logic can be used for setting-up the values:</p>
<pre><code class="lang-python"># Set the <span class="hljs-built_in">value</span> at <span class="hljs-built_in">second</span> <span class="hljs-built_in">row</span> <span class="hljs-built_in">and</span> third colu<span class="hljs-symbol">mn:</span>
sheet.iloc[<span class="hljs-number">1</span>,<span class="hljs-number">2</span>] = <span class="hljs-built_in">value</span>
# Set the <span class="hljs-built_in">value</span> by the name of the <span class="hljs-built_in">row</span>, <span class="hljs-built_in">column</span>
sheet.loc['super the label of <span class="hljs-built_in">row</span>', '<span class="hljs-built_in">even</span> better label of <span class="hljs-built_in">column</span>'] = <span class="hljs-built_in">value</span>
</code></pre>
<p>where the variable <code>value</code> can be either some constant (string, float or
created by the <code>fn</code> method described below) or the result of some
operations with cells:</p>
<pre><code class="lang-python">sheet.iloc[<span class="hljs-number">1</span>,<span class="hljs-number">2</span>] = sheet.iloc[<span class="hljs-number">1</span>,<span class="hljs-number">3</span>] + sheet.iloc[<span class="hljs-number">1</span>,<span class="hljs-number">4</span>]
</code></pre>
<p>In the case that you want to assign the result of some operation (or just
reference to another cell), make sure that it does not contains any reference
to itself (coordinates where you are assigning). It would not work
correctly otherwise.</p>
<h3 id="variables">Variables</h3>
<p>Variable represents an imaginary entity that can be used for computation if 
you want to refer to something that is common for the whole spreadsheet. 
Technically it is similar to variables in programming languages.</p>
<p>Variables are encapsulated in the property <code>var</code> of the class <code>Spreadsheet</code>. </p>
<p>It provides the following functionality:</p>
<ol>
<li><strong>Setting the variable</strong>, method <code>set_variable</code> with parameters: <code>name</code> 
(a lowercase alphanumeric string with underscores), <code>value</code> 
(number or string), and <code>description</code> (optional) that serves as a help text.</li>
<li><strong>Get the variable dictionary</strong>, property <code>variables_dict</code>, returns 
a dictionary with variable names as keys and variable values and descriptions
as values → following the logic: <code>{&#39;VARIABLE_NAME&#39;: {&#39;description&#39;:
&#39;String value or None&#39;, &#39;value&#39;: &#39;VALUE&#39;}}</code>.</li>
<li><strong>Check if the variable exists in a system</strong>, method <code>variable_exist</code> with
a parameter <code>name</code> representing the name of the variable. 
Return true if the variable exists, false otherwise.</li>
<li><strong>Get the variable as a Cell object</strong>, method <code>get_variable</code>, with
parameter <code>name</code> (required as positional only) that returns the variable as a
Cell object (for computations in a sheet).</li>
<li><strong>Check if there is any variable in the system</strong>: using the property <code>empty</code>
that returns true if there is no variable in the system, false otherwise. </li>
</ol>
<p>To get (and set similarly) the variable as a cell object, you can also use
the following approach with square brackets:</p>
<pre><code class="lang-python">sheet<span class="hljs-selector-class">.iloc</span>[<span class="hljs-selector-tag">i</span>, j] = sheet<span class="hljs-selector-class">.var</span>[<span class="hljs-string">'VARIABLE_NAME'</span>]
</code></pre>
<p>Same approach can be used for setting the value of variable:</p>
<pre><code class="lang-python">sheet<span class="hljs-selector-class">.var</span>[<span class="hljs-string">'VARIABLE_NAME'</span>] = some_value
</code></pre>
<p>Getting/setting the variables values should be done preferably by this logic.</p>
<p>For defining Excel format/style of the variable value, use the attribute
<code>excel_format</code> of the <code>var</code> property in the following logic:</p>
<pre><code class="lang-python">sheet<span class="hljs-selector-class">.var</span><span class="hljs-selector-class">.excel_format</span>[<span class="hljs-string">'VARIABLE_NAME'</span>] = {<span class="hljs-string">'num_format'</span>: <span class="hljs-string">'#,##0'</span>}
</code></pre>
<h4 id="example">Example</h4>
<p>Following example multiply some cell with value of
PI constant stored as a variable <code>pi</code>:</p>
<pre><code class="lang-python">sheet.set_variable(pi, <span class="hljs-number">3.14159265359</span>)
sheet.iloc[i,j] = sheet.var[<span class="hljs-string">'pi'</span>] * sheet.iloc[x,y]
</code></pre>
<h3 id="working-with-slices">Working with slices</h3>
<p>Similarly, like in NumPy or Pandas DataFrame, there is a possibility
how to work with slices (e. g. if you want to select a whole row, column
or set of rows and columns). Following code, select the third column:</p>
<pre><code class="lang-python"><span class="hljs-selector-tag">sheet</span><span class="hljs-selector-class">.iloc</span><span class="hljs-selector-attr">[:,2]</span>
</code></pre>
<p>On the other hand</p>
<pre><code class="lang-python">sheet<span class="hljs-selector-class">.loc</span>[:,<span class="hljs-string">'Handy column'</span>]
</code></pre>
<p>selects all the rows in the columns with the label <em>&#39;Handy column&#39;</em>. </p>
<p>You can again set the values in the slice to some constant, or the array
of constants, or to another cell, or to the result of some computation.</p>
<pre><code class="lang-python"><span class="hljs-keyword">sheet.iloc[:,2] </span>= constant  <span class="hljs-comment"># Constant (float, string)</span>
<span class="hljs-keyword">sheet.iloc[:,2] </span>= <span class="hljs-keyword">sheet.iloc[1,3] </span>+ <span class="hljs-keyword">sheet.iloc[1,4] </span> <span class="hljs-comment"># Computation result</span>
<span class="hljs-keyword">sheet.iloc[:,2] </span>= <span class="hljs-keyword">sheet.iloc[1,3] </span> <span class="hljs-comment"># Just a reference to a cell</span>
</code></pre>
<p>Technically the slice is the instance of <code>CellSlice</code> class.</p>
<p>There are two ways how to slice, either using <code>.loc</code> or <code>.iloc</code> attribute.
Where <code>iloc</code> uses integer position and <code>loc</code> uses label of the position
(as a string).</p>
<p>By default the right-most value is excluded when defining slices. If you want
to use right-most value indexing, use one of the methods described below.</p>
<h4 id="slicing-using-method-with-the-right-most-value-included-option-">Slicing using method (with the right-most value included option)</h4>
<p>Sometimes, it is quite helpful to use a slice that includes the right-most
value. There are two functions for this purpose:</p>
<ol>
<li><p><code>sheet.iloc.get_slice(ROW_INDEX, COLUMN_INDEX, include_right=[True/False])</code>:
This way is equivalent to the one presented above with square brackets <code>[]</code>.
The difference is the key-value attribute <code>include_right</code> that enables the
possibility of including the right-most value of the slice (default value is
False). If you want to use slice as your index, you need to pass some <code>slice</code>
object to one (or both) of the indices. For example: 
<code>sheet.iloc.get_slice(slice(0, 7), 3, include_right=True])</code> selects first nine
rows (because 8th row - right-most one - is included) from the fourth column
of the sheet <em>(remember, all is indexed from zero)</em>.</p>
</li>
<li><p><code>sheet.iloc.set_slice(ROW_INDEX, COLUMN_INDEX, VALUE, 
include_right=[True/False])</code>: this command set slice to <em>VALUE</em> in the similar
logic as when you call <code>get_slice</code> method (see the first point).</p>
</li>
</ol>
<p>There are again two possibilities, either to use <code>iloc</code> with integer position
or to use <code>loc</code> with labels.</p>
<h4 id="aggregate-functions">Aggregate functions</h4>
<p>The slice itself can be used for computations using aggregate functions.</p>
<ol>
<li><strong>Sum</strong>: return the sum of the cells in the slice. 
 For example: SUM(7, 8, 9) = 25.
 Available as the function <code>sum</code> called on the slice object.
 Usage: <code>sheet.iloc[i,j] = sheet.iloc[p:q,x:y].sum()</code></li>
<li><strong>Product</strong>: return the product of the cells in the slice. 
 For example: PROD(7, 8, 9) = 504.
 Available as the function <code>product</code> called on the slice object.
 Usage: <code>sheet.iloc[i,j] = sheet.iloc[p:q,x:y].product()</code></li>
<li><strong>Minimum</strong>: return the minimum of the cells in the slice. 
 For example: MIN(7, 8, 9) = 7.
 Available as the function <code>min</code> called on the slice object.
 Usage: <code>sheet.iloc[i,j] = sheet.iloc[p:q,x:y].min()</code></li>
<li><strong>Maximum</strong>: return the maximum of the cells in the slice. 
 For example: MAX(7, 8, 9) = 9.
 Available as the function <code>max</code> called on the slice object.
 Usage: <code>sheet.iloc[i,j] = sheet.iloc[p:q,x:y].max()</code></li>
<li><strong>Mean-average</strong>: return the arithmetic mean of the cells in the slice. 
 For example: MEAN(7, 8, 9) = 8.
 Available as the function <code>mean</code> and <code>average</code> called on the slice object.
 Usage: <code>sheet.iloc[i,j] = sheet.iloc[p:q,x:y].mean()</code> or 
 <code>sheet.iloc[i,j] = sheet.iloc[p:q,x:y].average()</code> </li>
<li><strong>Standard deviation</strong>: return the standard deviation of the cells in the
slice. 
 For example: STDEV(7, 8, 9) = 1.
 Available as the function <code>stdev</code> called on the slice object.
 Usage: <code>sheet.iloc[i,j] = sheet.iloc[p:q,x:y].stdev()</code></li>
<li><strong>Median</strong>: return the median of the cells in the slice. 
 For example: MEDIAN(7, 8, 9) = 8.
 Available as the function <code>median</code> called on the slice object.
 Usage: <code>sheet.iloc[i,j] = sheet.iloc[p:q,x:y].median()</code></li>
<li><strong>Count</strong>: return the number of the cells in the slice. 
 For example: COUNT(7, 8, 9) = 3.
 Available as the function <code>count</code> called on the slice object.
 Usage: <code>sheet.iloc[i,j] = sheet.iloc[p:q,x:y].count()</code></li>
<li><strong>IRR</strong>: return the Internal Rate of Return (IRR) of the cells in the slice. 
 For example: IRR(-100, 0, 0, 74) = -0.0955.
 Available as the function <code>irr</code> called on the slice object.
 Usage: <code>sheet.iloc[i,j] = sheet.iloc[p:q,x:y].irr()</code></li>
<li><strong>Match negative before positive</strong>: return the position of the last
negative number in a series of negative numbers in the row or column
series.
For example:
MNBP(-100, -90, -80, 5, -500) = 3 <em>(equals to position of the number -80)</em>.
Available as the function <code>match_negative_before_positive</code> called on the
slice object.
Usage:
<code>sheet.iloc[i,j] = sheet.iloc[p:q,x:y].match_negative_before_positive()</code></li>
</ol>
<p>Aggregate functions always return the cell with the result.</p>
<p>All aggregate functions have parameters:</p>
<ol>
<li><code>skip_none_cell (bool)</code>: If true, skips all the cells with <code>None</code> as
a value (and does not raise an exception), if false an exception is raised
if the slice contains a cell with <code>None</code> value (empty cell).</li>
</ol>
<h3 id="conditional">Conditional</h3>
<p>There is a support for the conditional statement (aka if-then-else statement).
Functionality is implemented in the property <code>fn</code> of the <code>Spreadsheet</code>
instance in the method <code>conditional</code>. It takes three parameters (positional)
in precisely this order:</p>
<ol>
<li><strong>Condition</strong>: the cell with a boolean value that is evaluated (typically
achieved using operators ==, !=, &gt;, &lt;, etc.).</li>
<li><strong>Consequent</strong>: the cell that is taken when the condition is evaluated as
true.</li>
<li><strong>Alternative</strong>:  the cell that is taken when the condition is evaluated as
false.</li>
</ol>
<p>All the parameters are the instance of <code>Cell</code> class.</p>
<h4 id="example-of-conditional">Example of conditional</h4>
<p>Consider the following example that compares whether two cells are equals,
if yes, it takes some value in a cell, if not, another value in the
different cell:</p>
<pre><code class="lang-python">sheet.iloc[i,j] = sheet.fn.conditional(
    # Condition <span class="hljs-built_in">is</span> the <span class="hljs-built_in">first</span> parameter:
    sheet.iloc[<span class="hljs-number">1</span>,<span class="hljs-number">2</span>] == sheet.iloc[<span class="hljs-number">2</span>,<span class="hljs-number">2</span>],
    # Consequent (value <span class="hljs-keyword">if</span> condition <span class="hljs-built_in">is</span> <span class="hljs-literal">true</span>) <span class="hljs-built_in">is</span> the <span class="hljs-built_in">second</span> parameter:
    sheet.iloc[<span class="hljs-number">3</span>,<span class="hljs-number">1</span>],
    # Alternative (value <span class="hljs-keyword">if</span> condition <span class="hljs-built_in">is</span> <span class="hljs-literal">false</span>) <span class="hljs-built_in">is</span> the <span class="hljs-built_in">third</span> parameter:
    sheet.iloc[<span class="hljs-number">4</span>,<span class="hljs-number">1</span>]
)
</code></pre>
<h3 id="raw-statement">Raw statement</h3>
<p>The raw statement represents the extreme way how to set-up value and
computation string of the cell. It should be used only to circumvent
issues with missing or defective functionality.</p>
<p>The raw statement is accessible using <code>fn</code> property of the Spreadsheet class
object.</p>
<p>The raw statement should never be used unless you really have to.</p>
<h4 id="example-of-raw-statement">Example of raw statement</h4>
<p>Consider that you need to compute an arccosine value of some cell:</p>
<pre><code class="lang-python">sheet.iloc[i,j] = sheet.fn.raw(
    # Value that should be used <span class="hljs-keyword">as</span> the result (<span class="hljs-keyword">as</span> a Cell <span class="hljs-keyword">instance</span>):
    sheet.fn.const(numpy.arccos(<span class="hljs-number">0.7</span>)),
    # Definition <span class="hljs-keyword">of</span> words <span class="hljs-keyword">in</span> each language:
    {
        <span class="hljs-string">'python_numpy'</span>: <span class="hljs-string">"numpy.arccos(0.7)"</span>,
        <span class="hljs-string">'excel'</span>: <span class="hljs-string">"ACOS(0.7)"</span>
        # Potentialy some other languages, like <span class="hljs-string">'native'</span>, etc.
    }
)
</code></pre>
<h3 id="offset-function">Offset function</h3>
<p>The offset function represents the possibility of reading the value
that is shifted by some number rows left, and some number of columns
down from some referential cells.</p>
<p>It is accessible from the Spreadsheet instance using <code>fn</code>
property and <code>offset</code> method. Parameters are following (only
positional, in exactly this order):</p>
<ul>
<li><strong>Reference cell</strong>: Reference cell from that the position is computed.</li>
<li><strong>Cell defining a number of rows to be skipped</strong>: How many rows (down)
should be skipped.</li>
<li><strong>Cell defining a number of columns to be skipped</strong>: How many columns (left)
should be skipped.</li>
</ul>
<h4 id="example-">Example:</h4>
<p>Following example assign the value of the cell that is on the third row and 
second column to the cell that is on the second row and second column.</p>
<pre><code class="lang-python">sheet<span class="hljs-selector-class">.iloc</span>[<span class="hljs-number">1</span>,<span class="hljs-number">1</span>] = sheet<span class="hljs-selector-class">.fn</span><span class="hljs-selector-class">.offset</span>(
    sheet<span class="hljs-selector-class">.iloc</span>[<span class="hljs-number">0</span>,<span class="hljs-number">0</span>], sheet<span class="hljs-selector-class">.fn</span><span class="hljs-selector-class">.const</span>(<span class="hljs-number">2</span>), sheet<span class="hljs-selector-class">.fn</span><span class="hljs-selector-class">.const</span>(<span class="hljs-number">1</span>)
)
</code></pre>
<h3 id="computations">Computations</h3>
<p>All operations have to be done using the objects of type Cell. </p>
<h4 id="constants">Constants</h4>
<p>If you want to use a constant value, you need to create an un-anchored cell
with this value. The easiest way of doing so is:</p>
<pre><code class="lang-python"># <span class="hljs-keyword">For</span> creating the Cell <span class="hljs-keyword">for</span> computation <span class="hljs-keyword">with</span> <span class="hljs-keyword">constant</span> value <span class="hljs-number">7</span>
constant_cell = sheet.fn.const(<span class="hljs-number">7</span>)
</code></pre>
<p>The value OPERAND bellow is always the reference to another cell in the
sheet or the constant created as just described.</p>
<h4 id="unary-operations">Unary operations</h4>
<p>There are the following unary operations (in the following the <code>OPERAND</code>
is the instance of the Cell class): </p>
<ol>
<li><strong>Ceiling function</strong>: returns the ceil of the input value.
 For example ceil(4.1) = 5.
 Available in the <code>fn</code> property of the <code>sheet</code> object.
 Usage: <code>sheet.iloc[i,j] = sheet.fn.ceil(OPERAND)</code></li>
<li><strong>Floor function</strong>: returns the floor of the input value.
 For example floor(4.1) = 4.
 Available in the <code>fn</code> property of the <code>sheet</code> object.
 Usage: <code>sheet.iloc[i,j] = sheet.fn.floor(OPERAND)</code></li>
<li><strong>Round function</strong>: returns the round of the input value.
 For example round(4.5) = 5.
 Available in the <code>fn</code> property of the <code>sheet</code> object.
 Usage: <code>sheet.iloc[i,j] = sheet.fn.round(OPERAND)</code></li>
<li><strong>Absolute value function</strong>: returns the absolute value of the input value.
 For example abs(-4.5) = 4.5.
 Available in the <code>fn</code> property of the <code>sheet</code> object.
 Usage: <code>sheet.iloc[i,j] = sheet.fn.abs(OPERAND)</code></li>
<li><strong>Square root function</strong>: returns the square root of the input value.
 For example sqrt(16) = 4.
 Available in the <code>fn</code> property of the <code>sheet</code> object.
 Usage: <code>sheet.iloc[i,j] = sheet.fn.sqrt(OPERAND)</code></li>
<li><strong>Logarithm function</strong>: returns the natural logarithm of the input value.
 For example ln(11) = 2.3978952728.
 Available in the <code>fn</code> property of the <code>sheet</code> object.
 Usage: <code>sheet.iloc[i,j] = sheet.fn.ln(OPERAND)</code></li>
<li><strong>Exponential function</strong>: returns the exponential of the input value.
 For example exp(1) = <em>e</em> power to 1 = 2.71828182846.
 Available in the <code>fn</code> property of the <code>sheet</code> object.
 Usage: <code>sheet.iloc[i,j] = sheet.fn.exp(OPERAND)</code></li>
<li><strong>Logical negation</strong>: returns the logical negation of the input value.
 For example neg(false) = true.
 Available as the overloaded operator <code>~</code>.
 Usage: <code>sheet.iloc[i,j] = ~OPERAND</code>.
 <em>Also available in the <code>fn</code> property of the <code>sheet</code> object.
 Usage: <code>sheet.iloc[i,j] = sheet.fn.neg(OPERAND)</code></em></li>
<li><strong>Signum function</strong>: returns the signum of the input value.
 For example sign(-4.5) = -1, sign(5) = 1, sign(0) = 0.
 Available in the <code>fn</code> property of the <code>sheet</code> object.
 Usage: <code>sheet.iloc[i,j] = sheet.fn.sign(OPERAND)</code></li>
</ol>
<p>All unary operators are defined in the <code>fn</code> property of the Spreadsheet
object (together with brackets, that works exactly the same - see bellow).</p>
<h4 id="binary-operations">Binary operations</h4>
<p>There are the following binary operations (in the following the <code>OPERAND_1</code>
and <code>OPERAND_2</code> are the instances of the Cell class):</p>
<ol>
<li><strong>Addition</strong>: return the sum of two numbers. 
 For example: 5 + 2 = 7.
 Available as the overloaded operator <code>+</code>.
 Usage: <code>sheet.iloc[i,j] = OPERAND_1 + OPERAND_2</code></li>
<li><strong>Subtraction</strong>: return the difference of two numbers. 
 For example: 5 - 2 = 3.
 Available as the overloaded operator <code>-</code>.
 Usage: <code>sheet.iloc[i,j] = OPERAND_1 - OPERAND_2</code></li>
<li><strong>Multiplication</strong>: return the product of two numbers. 
 For example: 5 <em> 2 = 10.
 Available as the overloaded operator `</em><code>.
 Usage:</code>sheet.iloc[i,j] = OPERAND_1 * OPERAND_2`</li>
<li><strong>Division</strong>: return the quotient of two numbers. 
 For example: 5 / 2 = 2.5.
 Available as the overloaded operator <code>/</code>.
 Usage: <code>sheet.iloc[i,j] = OPERAND_1 / OPERAND_2</code></li>
<li><strong>Exponentiation</strong>: return the power of two numbers. 
 For example: 5 <strong> 2 = 25.
 Available as the overloaded operator `</strong><code>.
 Usage:</code>sheet.iloc[i,j] = OPERAND_1 ** OPERAND_2`</li>
<li><strong>Logical equality</strong>: return true if inputs are equals, false otherwise. 
 For example: 5 = 2 &lt;=&gt; false.
 Available as the overloaded operator <code>==</code>.
 Usage: <code>sheet.iloc[i,j] = OPERAND_1 == OPERAND_2</code></li>
<li><strong>Logical inequality</strong>: return true if inputs are not equals,
false otherwise. 
 For example: 5 ≠ 2 &lt;=&gt; true.
 Available as the overloaded operator <code>!=</code>.
 Usage: <code>sheet.iloc[i,j] = OPERAND_1 != OPERAND_2</code></li>
<li><strong>Relational greater than operator</strong>: return true if the first operand is
greater than another operand, false otherwise. 
 For example: 5 &gt; 2 &lt;=&gt; true.
 Available as the overloaded operator <code>&gt;</code>.
 Usage: <code>sheet.iloc[i,j] = OPERAND_1 &gt; OPERAND_2</code></li>
<li><strong>Relational greater than or equal to operator</strong>: return true if the first
operand is greater than or equal to another operand, false otherwise. 
 For example: 5 ≥ 2 &lt;=&gt; true.
 Available as the overloaded operator <code>&gt;=</code>.
 Usage: <code>sheet.iloc[i,j] = OPERAND_1 &gt;= OPERAND_2</code></li>
<li><strong>Relational less than operator</strong>: return true if the first operand is
less than another operand, false otherwise. 
For example: 5 &lt; 2 &lt;=&gt; false.
Available as the overloaded operator <code>&lt;</code>.
Usage: <code>sheet.iloc[i,j] = OPERAND_1 &lt; OPERAND_2</code></li>
<li><strong>Relational less than or equal to operator</strong>: return true if the first
operand is less than or equal to another operand, false otherwise. 
For example: 5 ≤ 2 &lt;=&gt; false.
Available as the overloaded operator <code>&lt;=</code>.
Usage: <code>sheet.iloc[i,j] = OPERAND_1 &lt;= OPERAND_2</code></li>
<li><strong>Logical conjunction operator</strong>: return true if the first
operand is true and another operand is true, false otherwise. 
For example: true ∧ false &lt;=&gt; false.
Available as the overloaded operator <code>&amp;</code>.
Usage: <code>sheet.iloc[i,j] = OPERAND_1 &amp; OPERAND_2</code>.
<strong><em>BEWARE that operator <code>and</code> IS NOT OVERLOADED! Because it is not
technically possible.</em></strong></li>
<li><strong>Logical disjunction operator</strong>: return true if the first
operand is true or another operand is true, false otherwise. 
For example: true ∨ false &lt;=&gt; true.
Available as the overloaded operator <code>|</code>.
Usage: <code>sheet.iloc[i,j] = OPERAND_1 | OPERAND_2</code>.
<strong><em>BEWARE that operator <code>or</code> IS NOT OVERLOADED! Because it is not
technically possible.</em></strong></li>
<li><strong>Concatenate strings</strong>: return string concatenation of inputs.
For example: CONCATENATE(7, &quot;Hello&quot;) &lt;=&gt; &quot;7Hello&quot;.
Available as the overloaded operator <code>&lt;&lt;</code>.
Usage: <code>sheet.iloc[i,j] = OPERAND_1 &lt;&lt; OPERAND_2</code></li>
</ol>
<p>Operations can be chained in the string:</p>
<pre><code class="lang-python">sheet.iloc[i,j] = OPERA<span class="hljs-symbol">ND_1</span> + OPERA<span class="hljs-symbol">ND_2</span> * OPERA<span class="hljs-symbol">ND_3</span> ** OPERA<span class="hljs-symbol">ND_4</span>
</code></pre>
<p>The priority of the operators is the same as in normal mathematics. If
you need to modify priority, you need to use brackets, for example:</p>
<pre><code class="lang-python"><span class="hljs-keyword">sheet.iloc[i,j] </span>= <span class="hljs-keyword">sheet.fn.brackets(OPERAND_1 </span>+ OPERAND_2) \
    * OPERAND_3 ** OPERAND_4
</code></pre>
<h4 id="brackets-for-computation">Brackets for computation</h4>
<p>Brackets are technically speaking just another unary operator. They are
defined in the <code>fn</code> property. They can be used like:</p>
<pre><code class="lang-python"><span class="hljs-keyword">sheet.iloc[i,j] </span>= <span class="hljs-keyword">sheet.fn.brackets(OPERAND_1 </span>+ OPERAND_2) \ 
    * OPERAND_3 ** OPERAND_4
</code></pre>
<h4 id="example">Example</h4>
<p>For example</p>
<pre><code class="lang-python"># Equivalent <span class="hljs-keyword">of</span>: value at <span class="hljs-comment">[1,0]</span> * (value at <span class="hljs-comment">[2,1]</span> + value at <span class="hljs-comment">[3,1]</span>) * exp(9)
sheet.iloc<span class="hljs-comment">[0,0]</span> = sheet.iloc<span class="hljs-comment">[1,0]</span> * sheet.fn.brackets(
        sheet.iloc<span class="hljs-comment">[2,1]</span> + sheet.iloc<span class="hljs-comment">[3,1]</span>
    ) * sheet.fn.exp(sheet.fn.const(9))
</code></pre>
<h3 id="accessing-the-computed-values">Accessing the computed values</h3>
<p>You can access either to the actual numerical value of the cell or to the
word that is created in all the languages. The numerical value is accessible
using the <code>value</code> property, whereas the words are accessible using
the <code>parse</code> property (it returns a dictionary with languages as keys
and word as values).</p>
<pre><code class="lang-python"><span class="hljs-comment"># Access the value of the cell</span>
<span class="hljs-symbol">value_of_cell:</span> float = <span class="hljs-keyword">sheet.iloc[i, </span><span class="hljs-keyword">j].value
</span><span class="hljs-comment"># Access all the words in the cell</span>
<span class="hljs-symbol">word:</span> <span class="hljs-keyword">dict </span>= <span class="hljs-keyword">sheet.iloc[i, </span><span class="hljs-keyword">j].parse
</span><span class="hljs-comment"># Access the word in language 'lang'</span>
word_in_language_lang = word[<span class="hljs-string">'lang'</span>]
</code></pre>
<h3 id="exporting-the-results">Exporting the results</h3>
<p>There are various methods available for exporting the results. All these
methods can be used either to a whole sheet (instance of Spreadsheet)
or to any slice (CellSlice instance):</p>
<ol>
<li><strong>Excel format</strong>, method <code>to_excel</code>:
Export the sheet to the Excel-compatible file.</li>
<li><strong>Dictionary of values</strong>, method <code>to_dictionary</code>:
Export the sheet to the dictionary (<code>dict</code> type).</li>
<li><strong>JSON format</strong>, method <code>to_json</code>:
Export the sheet to the JSON format (serialize output of <code>to_dictionary</code>).</li>
<li><strong>2D array as a string</strong>, method: <code>to_string_of_values</code>:
Export values to the string that looks like Python array definition string.</li>
<li><strong>CSV</strong>, method <code>to_csv</code>:
Export the values to the CSV compatible string (that can be saved to the file)</li>
<li><strong>Markdown (MD)</strong>, method <code>to_markdown</code>:
Export the values to MD (Markdown) file format string.
Defined as a table.</li>
<li><strong>NumPy ndarray</strong>, method <code>to_numpy</code>:
Export the sheet as a <code>numpy.ndarray</code> object.</li>
<li><strong>Python 2D list</strong>, method <code>to_2d_list</code>: 
Export values 2 dimensional Python array (list of the list of the values).</li>
<li><strong>HTML table</strong>, method <code>to_html_table</code>:
Export values to HTML table.</li>
</ol>
<h4 id="description-field">Description field</h4>
<p>There is a possibility to add a description to a cell in the sheet
(or to the whole slice of the sheet). It can be done using the property
<code>description</code> on the cell or slice object. It should be done just before
the export is done (together with defining Excel styles, see below)
because once you rewrite the value of the cell on a given location,
the description is lost.</p>
<p>Example of using the description field:</p>
<pre><code class="lang-python"><span class="hljs-comment"># Setting the description of a single cell</span>
<span class="hljs-keyword">sheet.iloc[i, </span><span class="hljs-keyword">j].description </span>= <span class="hljs-string">"Some text describing a cell"</span>
<span class="hljs-comment"># Seting the description to a slice (propagate its value to each cell)</span>
<span class="hljs-keyword">sheet.iloc[i:j, </span>k:l].description = <span class="hljs-string">"Text describing each cell in the slice"</span>
</code></pre>
<h4 id="exporting-to-excel">Exporting to Excel</h4>
<p>It can be done using the interface:</p>
<pre><code class="lang-python">sheet.to_excel(
    file_path: <span class="hljs-built_in">str</span>,
    /, *,
    sheet_name: <span class="hljs-built_in">str</span> = <span class="hljs-string">"Results"</span>,
    spaces_replacement: <span class="hljs-built_in">str</span> = <span class="hljs-string">' '</span>,
    label_row_format: dict = {<span class="hljs-string">'bold'</span>: <span class="hljs-literal">True</span>},
    label_column_format: dict = {<span class="hljs-string">'bold'</span>: <span class="hljs-literal">True</span>},
    variables_sheet_name: Optional[<span class="hljs-built_in">str</span>] = None,
    variables_sheet_header: Dict[<span class="hljs-built_in">str</span>, <span class="hljs-built_in">str</span>] = MappingProxyType(
    {
        <span class="hljs-string">"name"</span>: <span class="hljs-string">"Name"</span>,
        <span class="hljs-string">"value"</span>: <span class="hljs-string">"Value"</span>,
        <span class="hljs-string">"description"</span>: <span class="hljs-string">"Description"</span>
    }),
    values_only: bool = <span class="hljs-literal">False</span>,
    skipped_label_replacement: <span class="hljs-built_in">str</span> = <span class="hljs-string">''</span>,
    row_height: <span class="hljs-built_in">List</span>[float] = [],
    column_width: <span class="hljs-built_in">List</span>[float] = [],
    top_left_corner_text: <span class="hljs-built_in">str</span> = <span class="hljs-string">""</span>
)
</code></pre>
<p>The only required argument is the path to the destination file (positional
only parameter). Other parameters are passed as keywords (non-positional only). </p>
<ul>
<li><code>file_path (str)</code>: Path to the target .xlsx file. (<strong>REQUIRED</strong>, only
positional)</li>
<li><code>sheet_name (str)</code>: The name of the sheet inside the file.</li>
<li><code>spaces_replacement (str)</code>: All the spaces in the rows and columns
descriptions (labels) are replaced with this string.</li>
<li><code>label_row_format (dict)</code>: Excel styles for the label of rows,
documentation: <a href="https://xlsxwriter.readthedocs.io/format.html">https://xlsxwriter.readthedocs.io/format.html</a></li>
<li><code>label_column_format (dict)</code>: Excel styles for the label of columns,
documentation: <a href="https://xlsxwriter.readthedocs.io/format.html">https://xlsxwriter.readthedocs.io/format.html</a></li>
<li><code>variables_sheet_name (Optional[str])</code>: If set, creates the new
sheet with variables and their description and possibility
to set them up (directly from the sheet).</li>
<li><code>variables_sheet_header (Dict[str, str])</code>: Define the labels (header)
for the sheet with variables (first row in the sheet). Dictionary should look
like: <code>{&quot;name&quot;: &quot;Name&quot;, &quot;value&quot;: &quot;Value&quot;, &quot;description&quot;: &quot;Description&quot;}</code>.</li>
<li><code>values_only (bool)</code>: If true, only values (and not formulas) are
exported.</li>
<li><code>skipped_label_replacement (str)</code>: Replacement for the SkippedLabel
instances.</li>
<li><code>row_height (List[float])</code>: List of row heights, or empty for the
default height (or <code>None</code> for default height in the series).
If row labels are included, there is a label row height on the
first position in array.</li>
<li><code>column_width (List[float])</code>: List of column widths, or empty for the
default widths (or <code>None</code> for the default width in the series).
If column labels are included, there is a label column width
on the first position in array.</li>
<li><code>top_left_corner_text (str)</code>: Text in the top left corner. Apply
only when the row and column labels are included.</li>
</ul>
<h5 id="setting-the-format-style-for-excel-cells">Setting the format/style for Excel cells</h5>
<p>There is a possibility to set the style/format of each cell in the grid
or the slice of the gird using property <code>excel_format</code>. Style assignment
should be done just before the export to the file because each new
assignment of values to the cell overrides its style. Format/style can
be set for both slice and single value. </p>
<p>Example of setting Excel format/style for cells and slices:</p>
<pre><code class="lang-python"><span class="hljs-comment"># Set the format of the cell on the position [i, j] (use bold value)</span>
sheet.iloc[i, j].excel_format = {<span class="hljs-string">'bold'</span>: <span class="hljs-keyword">True</span>}
<span class="hljs-comment"># Set the format of the cell slice (use bold value and red color)</span>
sheet.iloc[i:j, k:l].excel_format = {<span class="hljs-string">'bold'</span>: <span class="hljs-keyword">True</span>, <span class="hljs-string">'color'</span>: <span class="hljs-string">'red'</span>}
</code></pre>
<h5 id="appending-to-existing-excel-file">Appending to existing Excel file</h5>
<p>Appending to existing Excel (<code>.xlsx</code>) format <strong>is currently not supported</strong> due
to the missing functionality of the package XlsxWriter on which this
library relies.</p>
<h4 id="exporting-to-the-dictionary-and-json-">Exporting to the dictionary (and JSON)</h4>
<p>It can be done using the interface:</p>
<pre><code class="lang-python"><span class="hljs-selector-tag">sheet</span><span class="hljs-selector-class">.to_dictionary</span>(<span class="hljs-attribute">languages</span>: List[str] = None,
                    <span class="hljs-attribute">use_language_for_description</span>: Optional[str] = None, 
                    /, *, 
                    <span class="hljs-attribute">by_row</span>: bool = True,
                    <span class="hljs-attribute">languages_pseudonyms</span>: List[str] = None,
                    <span class="hljs-attribute">spaces_replacement</span>: str = <span class="hljs-string">' '</span>,
                    <span class="hljs-attribute">skip_nan_cell</span>: bool = False,
                    <span class="hljs-attribute">nan_replacement</span>: object = None,
                    <span class="hljs-attribute">append_dict</span>: dict = {})
</code></pre>
<p><strong>Parameters are (all optional):</strong></p>
<p><em>Positional only:</em></p>
<ul>
<li><code>languages (List[str])</code>: List of languages that should be exported.</li>
<li><code>use_language_for_description (Optional[str])</code>: If set-up (using the language
name), description field is set to be either the description value 
(if defined) or the value of this language. </li>
</ul>
<p><em>Key-value only:</em></p>
<ul>
<li><code>by_row (bool)</code>: If True, rows are the first indices and columns are the
second in the order. If False it is vice-versa.</li>
<li><code>languages_pseudonyms (List[str])</code>: Rename languages to the strings inside
this list.</li>
<li><code>spaces_replacement (str)</code>: All the spaces in the rows and columns
descriptions (labels) are replaced with this string.</li>
<li><code>skip_nan_cell (bool)</code>: If true, <code>None</code> (NaN, empty cells) values are
skipped, default value is false (NaN values are included).</li>
<li><code>nan_replacement (object)</code>: Replacement for the <code>None</code> (NaN) value.</li>
<li><code>error_replacement (object)</code>: Replacement for the error value.</li>
<li><code>append_dict (dict)</code>: Append this dictionary to output.</li>
<li><code>generate_schema (bool)</code>: If true, returns the JSON schema.</li>
</ul>
<p>All the rows and columns with labels that are instances of SkippedLabel are
entirely skipped. </p>
<p><strong>The return value is:</strong> </p>
<p>Dictionary with keys: 1. column/row, 2. row/column, 3. language or
language pseudonym or &#39;value&#39; keyword for values -&gt; value as a value or
as a cell building string.</p>
<h5 id="exporting-to-json">Exporting to JSON</h5>
<p>Exporting to JSON string is available using <code>to_json</code> method with exactly the
same interface. The return value is the string.</p>
<p>The reason why this method is separate is because of some values inserted
from NumPy arrays cannot be serialized using native serializer.</p>
<p>To get JSON schema you can use either <code>generate_schema (bool)</code> parameter or
directly use static method <code>generate_json_schema</code> of the <code>Spreadsheet</code> class.</p>
<h5 id="output-example">Output example</h5>
<p>Output of the JSON format</p>
<pre><code class="lang-json">{
   <span class="hljs-attr">"table"</span>:{
      <span class="hljs-attr">"data"</span>:{
         <span class="hljs-attr">"rows"</span>:{
            <span class="hljs-attr">"R_0"</span>:{
               <span class="hljs-attr">"columns"</span>:{
                  <span class="hljs-attr">"NL_C_0"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"1"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"1"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"1"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">1</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor0,0"</span>
                  },
                  <span class="hljs-attr">"NL_C_1"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"2"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"2"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"2"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">2</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor0,1"</span>
                  },
                  <span class="hljs-attr">"NL_C_2"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"3"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"3"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"3"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">3</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor0,2"</span>
                  },
                  <span class="hljs-attr">"NL_C_3"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"4"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"4"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"4"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">4</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor0,3"</span>
                  }
               }
            },
            <span class="hljs-attr">"R_1"</span>:{
               <span class="hljs-attr">"columns"</span>:{
                  <span class="hljs-attr">"NL_C_0"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"5"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"5"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"5"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">5</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor1,0"</span>
                  },
                  <span class="hljs-attr">"NL_C_1"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"6"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"6"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"6"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">6</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor1,1"</span>
                  },
                  <span class="hljs-attr">"NL_C_2"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"7"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"7"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"7"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">7</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor1,2"</span>
                  },
                  <span class="hljs-attr">"NL_C_3"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"8"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"8"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"8"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">8</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor1,3"</span>
                  }
               }
            },
            <span class="hljs-attr">"R_2"</span>:{
               <span class="hljs-attr">"columns"</span>:{
                  <span class="hljs-attr">"NL_C_0"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"9"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"9"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"9"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">9</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor2,0"</span>
                  },
                  <span class="hljs-attr">"NL_C_1"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"10"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"10"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"10"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">10</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor2,1"</span>
                  },
                  <span class="hljs-attr">"NL_C_2"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"11"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"11"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"11"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">11</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor2,2"</span>
                  },
                  <span class="hljs-attr">"NL_C_3"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"12"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"12"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"12"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">12</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor2,3"</span>
                  }
               }
            },
            <span class="hljs-attr">"R_3"</span>:{
               <span class="hljs-attr">"columns"</span>:{
                  <span class="hljs-attr">"NL_C_0"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"13"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"13"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"13"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">13</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor3,0"</span>
                  },
                  <span class="hljs-attr">"NL_C_1"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"14"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"14"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"14"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">14</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor3,1"</span>
                  },
                  <span class="hljs-attr">"NL_C_2"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"15"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"15"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"15"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">15</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor3,2"</span>
                  },
                  <span class="hljs-attr">"NL_C_3"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"16"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"16"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"16"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">16</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor3,3"</span>
                  }
               }
            },
            <span class="hljs-attr">"R_4"</span>:{
               <span class="hljs-attr">"columns"</span>:{
                  <span class="hljs-attr">"NL_C_0"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"17"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"17"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"17"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">17</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor4,0"</span>
                  },
                  <span class="hljs-attr">"NL_C_1"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"18"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"18"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"18"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">18</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor4,1"</span>
                  },
                  <span class="hljs-attr">"NL_C_2"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"19"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"19"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"19"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">19</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor4,2"</span>
                  },
                  <span class="hljs-attr">"NL_C_3"</span>:{
                     <span class="hljs-attr">"excel"</span>:<span class="hljs-string">"20"</span>,
                     <span class="hljs-attr">"python_numpy"</span>:<span class="hljs-string">"20"</span>,
                     <span class="hljs-attr">"native"</span>:<span class="hljs-string">"20"</span>,
                     <span class="hljs-attr">"value"</span>:<span class="hljs-number">20</span>,
                     <span class="hljs-attr">"description"</span>:<span class="hljs-string">"DescFor4,3"</span>
                  }
               }
            }
         }
      },
      <span class="hljs-attr">"variables"</span>:{

      },
      <span class="hljs-attr">"rows"</span>:[
         {
            <span class="hljs-attr">"name"</span>:<span class="hljs-string">"R_0"</span>,
            <span class="hljs-attr">"description"</span>:<span class="hljs-string">"HT_R_0"</span>
         },
         {
            <span class="hljs-attr">"name"</span>:<span class="hljs-string">"R_1"</span>,
            <span class="hljs-attr">"description"</span>:<span class="hljs-string">"HT_R_1"</span>
         },
         {
            <span class="hljs-attr">"name"</span>:<span class="hljs-string">"R_2"</span>,
            <span class="hljs-attr">"description"</span>:<span class="hljs-string">"HT_R_2"</span>
         },
         {
            <span class="hljs-attr">"name"</span>:<span class="hljs-string">"R_3"</span>,
            <span class="hljs-attr">"description"</span>:<span class="hljs-string">"HT_R_3"</span>
         },
         {
            <span class="hljs-attr">"name"</span>:<span class="hljs-string">"R_4"</span>,
            <span class="hljs-attr">"description"</span>:<span class="hljs-string">"HT_R_4"</span>
         }
      ],
      <span class="hljs-attr">"columns"</span>:[
         {
            <span class="hljs-attr">"name"</span>:<span class="hljs-string">"NL_C_0"</span>,
            <span class="hljs-attr">"description"</span>:<span class="hljs-string">"HT_C_0"</span>
         },
         {
            <span class="hljs-attr">"name"</span>:<span class="hljs-string">"NL_C_1"</span>,
            <span class="hljs-attr">"description"</span>:<span class="hljs-string">"HT_C_1"</span>
         },
         {
            <span class="hljs-attr">"name"</span>:<span class="hljs-string">"NL_C_2"</span>,
            <span class="hljs-attr">"description"</span>:<span class="hljs-string">"HT_C_2"</span>
         },
         {
            <span class="hljs-attr">"name"</span>:<span class="hljs-string">"NL_C_3"</span>,
            <span class="hljs-attr">"description"</span>:<span class="hljs-string">"HT_C_3"</span>
         }
      ]
   }
}
</code></pre>
<h4 id="exporting-to-the-csv">Exporting to the CSV</h4>
<p>It can be done using the interface:</p>
<pre><code class="lang-python">sheet.to_csv(*,
    <span class="hljs-built_in">language</span>: Optional[<span class="hljs-built_in">str</span>] = None,
    spaces_replacement: <span class="hljs-built_in">str</span> = <span class="hljs-string">' '</span>,
    top_left_corner_text: <span class="hljs-built_in">str</span> = <span class="hljs-string">"Sheet"</span>,
    sep: <span class="hljs-built_in">str</span> = <span class="hljs-string">','</span>,
    line_terminator: <span class="hljs-built_in">str</span> = <span class="hljs-string">'\n'</span>,
    na_rep: <span class="hljs-built_in">str</span> = <span class="hljs-string">''</span>,
    skip_labels: bool = <span class="hljs-literal">False</span>,
    skipped_label_replacement: <span class="hljs-built_in">str</span> = <span class="hljs-string">''</span>
) -&gt; <span class="hljs-built_in">str</span>
</code></pre>
<p>Parameters are (all optional and key-value only):</p>
<ul>
<li><code>language (Optional[str])</code>: If set-up, export the word in this
language in each cell instead of values.</li>
<li><code>spaces_replacement (str)</code>: All the spaces in the rows and columns
descriptions (labels) are replaced with this string.</li>
<li><code>top_left_corner_text (str)</code>: Text in the top left corner.</li>
<li><code>sep (str)</code>: Separator of values in a row.</li>
<li><code>line_terminator (str)</code>: Ending sequence (character) of a row.</li>
<li><code>na_rep (str)</code>: Replacement for the missing data.</li>
<li><code>skip_labels (bool)</code>: If true, first row and column with labels is
skipped</li>
<li><code>skipped_label_replacement (str)</code>: Replacement for the SkippedLabel
instances.</li>
</ul>
<p><strong>The return value is:</strong> </p>
<p>CSV of the values as a string.</p>
<h5 id="output-example">Output example</h5>
<pre><code class="lang-text">Sheet,NL_C_0,NL_C_1,NL_C_2,NL_C_3
R_0,<span class="hljs-number">1</span>,<span class="hljs-number">2</span>,<span class="hljs-number">3</span>,<span class="hljs-number">4</span>
R_1,<span class="hljs-number">5</span>,<span class="hljs-number">6</span>,<span class="hljs-number">7</span>,<span class="hljs-number">8</span>
R_2,<span class="hljs-number">9</span>,<span class="hljs-number">10</span>,<span class="hljs-number">11</span>,<span class="hljs-number">12</span>
R_3,<span class="hljs-number">13</span>,<span class="hljs-number">14</span>,<span class="hljs-number">15</span>,<span class="hljs-number">16</span>
R_4,<span class="hljs-number">17</span>,<span class="hljs-number">18</span>,<span class="hljs-number">19</span>,<span class="hljs-number">20</span>
</code></pre>
<h4 id="exporting-to-markdown-md-format">Exporting to Markdown (MD) format</h4>
<p>It can be done using the interface:</p>
<pre><code class="lang-python"><span class="hljs-selector-tag">sheet</span><span class="hljs-selector-class">.to_markdown</span>(*,
    <span class="hljs-attribute">language</span>: Optional[str] = None,
    <span class="hljs-attribute">spaces_replacement</span>: str = <span class="hljs-string">' '</span>,
    <span class="hljs-attribute">top_left_corner_text</span>: str = <span class="hljs-string">"Sheet"</span>,
    <span class="hljs-attribute">na_rep</span>: str = <span class="hljs-string">''</span>,
    <span class="hljs-attribute">skip_labels</span>: bool = False,
    <span class="hljs-attribute">skipped_label_replacement</span>: str = <span class="hljs-string">''</span>
)
</code></pre>
<p>Parameters are (all optional, all key-value only):</p>
<ul>
<li><code>language (Optional[str])</code>: If set-up, export the word in this
language in each cell instead of values.</li>
<li><code>spaces_replacement (str)</code>: All the spaces in the rows and columns
descriptions (labels) are replaced with this string.</li>
<li><code>top_left_corner_text (str)</code>: Text in the top left corner.</li>
<li><code>na_rep (str)</code>: Replacement for the missing data.</li>
<li><code>skip_labels (bool)</code>: If true, first row and column with labels is
skipped</li>
<li><code>skipped_label_replacement (str)</code>: Replacement for the SkippedLabel
instances.</li>
</ul>
<p><strong>The return value is:</strong> </p>
<p>Markdown (MD) compatible table of the values as a string.</p>
<h5 id="output-example">Output example</h5>
<pre><code class="lang-markdown">|<span class="hljs-string"> Sheet </span>|<span class="hljs-string">*NL_C_0* </span>|<span class="hljs-string"> *NL_C_1* </span>|<span class="hljs-string"> *NL_C_2* </span>|<span class="hljs-string"> *NL_C_3* </span>|
|<span class="hljs-string">----</span>|<span class="hljs-string">----</span>|<span class="hljs-string">----</span>|<span class="hljs-string">----</span>|<span class="hljs-string">----</span>|
|<span class="hljs-string"> *R_0* </span>|<span class="hljs-string"> 1 </span>|<span class="hljs-string"> 2 </span>|<span class="hljs-string"> 3 </span>|<span class="hljs-string"> 4 </span>|
|<span class="hljs-string"> *R_1* </span>|<span class="hljs-string"> 5 </span>|<span class="hljs-string"> 6 </span>|<span class="hljs-string"> 7 </span>|<span class="hljs-string"> 8 </span>|
|<span class="hljs-string"> *R_2* </span>|<span class="hljs-string"> 9 </span>|<span class="hljs-string"> 10 </span>|<span class="hljs-string"> 11 </span>|<span class="hljs-string"> 12 </span>|
|<span class="hljs-string"> *R_3* </span>|<span class="hljs-string"> 13 </span>|<span class="hljs-string"> 14 </span>|<span class="hljs-string"> 15 </span>|<span class="hljs-string"> 16 </span>|
|<span class="hljs-string"> *R_4* </span>|<span class="hljs-string"> 17 </span>|<span class="hljs-string"> 18 </span>|<span class="hljs-string"> 19 </span>|<span class="hljs-string"> 20 </span>|
</code></pre>
<h4 id="exporting-to-html-table-format">Exporting to HTML table format</h4>
<p>It can be done using the interface:</p>
<pre><code class="lang-python"><span class="hljs-selector-tag">sheet</span><span class="hljs-selector-class">.to_html_table</span>(*,
    <span class="hljs-attribute">spaces_replacement</span>: str = <span class="hljs-string">' '</span>,
    <span class="hljs-attribute">top_left_corner_text</span>: str = <span class="hljs-string">"Sheet"</span>,
    <span class="hljs-attribute">na_rep</span>: str = <span class="hljs-string">''</span>,
    <span class="hljs-attribute">language_for_description</span>: str = None,
    <span class="hljs-attribute">skip_labels</span>: bool = False,
    <span class="hljs-attribute">skipped_label_replacement</span>: str = <span class="hljs-string">''</span>
)
</code></pre>
<p>Parameters are (all optional, all key-value only):</p>
<ul>
<li><code>spaces_replacement (str)</code>: All the spaces in the rows and columns
descriptions (labels) are replaced with this string.</li>
<li><code>top_left_corner_text (str)</code>: Text in the top left corner.</li>
<li><code>na_rep (str)</code>: Replacement for the missing data.</li>
<li><code>language_for_description (str)</code>: If not <code>None</code>, the description
of each computational cell is inserted as word of this language
(if the property description is not set).</li>
<li><code>skip_labels (bool)</code>: If true, first row and column with labels is
skipped</li>
<li><code>skipped_label_replacement (str)</code>: Replacement for the SkippedLabel
instances.</li>
</ul>
<p><strong>The return value is:</strong> </p>
<p>HTML table of the values as a string. Table is usable mainly for debugging
purposes.</p>
<h5 id="output-example">Output example</h5>
<pre><code class="lang-html"><span class="hljs-tag">&lt;<span class="hljs-name">table</span>&gt;</span>
   <span class="hljs-tag">&lt;<span class="hljs-name">tr</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">th</span>&gt;</span>Sheet<span class="hljs-tag">&lt;/<span class="hljs-name">th</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">th</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"HT_C_0"</span>&gt;</span>NL_C_0<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">th</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">th</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"HT_C_1"</span>&gt;</span>NL_C_1<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">th</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">th</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"HT_C_2"</span>&gt;</span>NL_C_2<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">th</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">th</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"HT_C_3"</span>&gt;</span>NL_C_3<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">th</span>&gt;</span>
   <span class="hljs-tag">&lt;/<span class="hljs-name">tr</span>&gt;</span>
   <span class="hljs-tag">&lt;<span class="hljs-name">tr</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"HT_R_0"</span>&gt;</span>R_0<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor0,0"</span>&gt;</span>1<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor0,1"</span>&gt;</span>2<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor0,2"</span>&gt;</span>3<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor0,3"</span>&gt;</span>4<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
   <span class="hljs-tag">&lt;/<span class="hljs-name">tr</span>&gt;</span>
   <span class="hljs-tag">&lt;<span class="hljs-name">tr</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"HT_R_1"</span>&gt;</span>R_1<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor1,0"</span>&gt;</span>5<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor1,1"</span>&gt;</span>6<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor1,2"</span>&gt;</span>7<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor1,3"</span>&gt;</span>8<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
   <span class="hljs-tag">&lt;/<span class="hljs-name">tr</span>&gt;</span>
   <span class="hljs-tag">&lt;<span class="hljs-name">tr</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"HT_R_2"</span>&gt;</span>R_2<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor2,0"</span>&gt;</span>9<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor2,1"</span>&gt;</span>10<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor2,2"</span>&gt;</span>11<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor2,3"</span>&gt;</span>12<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
   <span class="hljs-tag">&lt;/<span class="hljs-name">tr</span>&gt;</span>
   <span class="hljs-tag">&lt;<span class="hljs-name">tr</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"HT_R_3"</span>&gt;</span>R_3<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor3,0"</span>&gt;</span>13<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor3,1"</span>&gt;</span>14<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor3,2"</span>&gt;</span>15<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor3,3"</span>&gt;</span>16<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
   <span class="hljs-tag">&lt;/<span class="hljs-name">tr</span>&gt;</span>
   <span class="hljs-tag">&lt;<span class="hljs-name">tr</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"HT_R_4"</span>&gt;</span>R_4<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor4,0"</span>&gt;</span>17<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor4,1"</span>&gt;</span>18<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor4,2"</span>&gt;</span>19<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
      <span class="hljs-tag">&lt;<span class="hljs-name">td</span>&gt;</span><span class="hljs-tag">&lt;<span class="hljs-name">a</span> <span class="hljs-attr">href</span>=<span class="hljs-string">"javascript:;"</span>  <span class="hljs-attr">title</span>=<span class="hljs-string">"DescFor4,3"</span>&gt;</span>20<span class="hljs-tag">&lt;/<span class="hljs-name">a</span>&gt;</span><span class="hljs-tag">&lt;/<span class="hljs-name">td</span>&gt;</span>
   <span class="hljs-tag">&lt;/<span class="hljs-name">tr</span>&gt;</span>
<span class="hljs-tag">&lt;/<span class="hljs-name">table</span>&gt;</span>
</code></pre>
<h2 id="remarks-and-definitions">Remarks and definitions</h2>
<ul>
<li><strong>Anchored cell</strong> is a cell that is located in the sheet and can be
accessed using position.</li>
<li><strong>Un-anchored cell</strong> is a cell that is the result of some computation or
a constant defined by the user for some computation (and does not have
any position in the sheet grid yet).</li>
</ul>
<p><strong>Example:</strong></p>
<pre><code class="lang-python">anchored_cell = sheet.iloc[<span class="hljs-number">4</span>,<span class="hljs-number">2</span>]
unanchored_cell_1 = sheet.iloc[<span class="hljs-number">4</span>,<span class="hljs-number">2</span>] * sheet.iloc[<span class="hljs-number">5</span>,<span class="hljs-number">2</span>]
unanchored_cell_2 = sheet.fn.const(<span class="hljs-number">9</span>)
</code></pre>
<h2 id="software-user-manual-sum-how-to-use-it-">Software User Manual (SUM), how to use it?</h2>
<h3 id="installation">Installation</h3>
<p>To install the most actual package, use the command:</p>
<pre><code class="lang-commandline">git clone http<span class="hljs-variable">s:</span>//github.<span class="hljs-keyword">com</span>/david-salac/Portable-spreadsheet-generator
<span class="hljs-keyword">cd</span> Portable-spreadsheet-generator/
<span class="hljs-keyword">python</span> setup.<span class="hljs-keyword">py</span> install
</code></pre>
<p>or simply install using PIP:</p>
<pre><code class="lang-commandline">pip <span class="hljs-keyword">install</span> portable-spreadsheet
</code></pre>
<h4 id="running-of-the-unit-tests">Running of the unit-tests</h4>
<p>For running package unit-tests, use command:</p>
<pre><code class="lang-commandline">python setup<span class="hljs-selector-class">.py</span> test
</code></pre>
<p>In order to run package unit-tests you need to clone package first.</p>
<h3 id="demo">Demo</h3>
<p>The following demo contains a simple example with aggregations.</p>
<pre><code class="lang-python">import portable_spreadsheet <span class="hljs-keyword">as</span> ps
import numpy <span class="hljs-keyword">as</span> np

<span class="hljs-comment"># This is a simple demo that represents the possibilities of the package</span>
<span class="hljs-comment">#   The purpose of this demo is to create a class rooms and monitor students</span>

sheet = ps.Spreadsheet.create_new_sheet(
    <span class="hljs-comment"># Size of the table (rows, columns):</span>
    <span class="hljs-number">24</span>, <span class="hljs-number">8</span>,
    rows_labels=[<span class="hljs-string">'Adam'</span>, <span class="hljs-string">'Oliver'</span>, <span class="hljs-string">'Harry'</span>, <span class="hljs-string">'George'</span>, <span class="hljs-string">'John'</span>, <span class="hljs-string">'Jack'</span>, <span class="hljs-string">'Jacob'</span>,
                 <span class="hljs-string">'Leo'</span>, <span class="hljs-string">'Oscar'</span>, <span class="hljs-string">'Charlie'</span>, <span class="hljs-string">'Peter'</span>, <span class="hljs-string">'Olivia'</span>, <span class="hljs-string">'Amelia'</span>,
                 <span class="hljs-string">'Isla'</span>, <span class="hljs-string">'Ava'</span>, <span class="hljs-string">'Emily'</span>, <span class="hljs-string">'Isabella'</span>, <span class="hljs-string">'Mia'</span>, <span class="hljs-string">'Poppy'</span>,
                 <span class="hljs-string">'Ella'</span>, <span class="hljs-string">'Lily'</span>, <span class="hljs-string">'Average of all'</span>, <span class="hljs-string">'Average of boys'</span>,
                 <span class="hljs-string">'Average of girls'</span>],
    columns_labels=[<span class="hljs-string">'Biology'</span>, <span class="hljs-string">'Physics'</span>, <span class="hljs-string">'Math'</span>, <span class="hljs-string">'English'</span>, <span class="hljs-string">'French'</span>,
                    <span class="hljs-string">'Best performance'</span>, <span class="hljs-string">'Worst performance'</span>, <span class="hljs-string">'Mean'</span>],
    columns_help_text=[
        <span class="hljs-string">'Annual performance'</span>, <span class="hljs-string">'Annual performance'</span>, <span class="hljs-string">'Annual performance'</span>,
        <span class="hljs-string">'Annual performance'</span>, <span class="hljs-string">'Annual performance'</span>,
        <span class="hljs-string">'Best performance of all subjects'</span>,
        <span class="hljs-string">'Worst performance of all subjects'</span>,
        <span class="hljs-string">'Mean performance of all subjects'</span>,
    ]
)

<span class="hljs-comment"># === Insert some percentiles to students performance: ===</span>
<span class="hljs-comment"># A) In this case insert random values in the first row to the 3rd row from the</span>
<span class="hljs-comment">#   end, and in the first column.</span>
sheet.iloc[:<span class="hljs-number">-3</span>, <span class="hljs-number">0</span>] = np.<span class="hljs-built_in">random</span>.<span class="hljs-built_in">random</span>(<span class="hljs-number">21</span>) * <span class="hljs-number">100</span>
<span class="hljs-comment"># B) Same can be achieved using the label indices:</span>
sheet.loc[<span class="hljs-string">"Adam"</span>:<span class="hljs-string">'Average of all'</span>, <span class="hljs-string">'Physics'</span>] = np.<span class="hljs-built_in">random</span>.<span class="hljs-built_in">random</span>(<span class="hljs-number">21</span>) * <span class="hljs-number">100</span>
<span class="hljs-comment"># C) Or by using the cell by cell approach:</span>
<span class="hljs-keyword">for</span> row_idx <span class="hljs-keyword">in</span> range(<span class="hljs-number">21</span>):
    <span class="hljs-comment"># I) Again by the simple integer index</span>
    sheet.iloc[row_idx, <span class="hljs-number">2</span>] = np.<span class="hljs-built_in">random</span>.<span class="hljs-built_in">random</span>() * <span class="hljs-number">100</span>
    <span class="hljs-comment"># II) Or by the label</span>
    row_label: str = sheet.cell_indices.rows_labels[row_idx]
    sheet.loc[row_label, <span class="hljs-string">'English'</span>] = np.<span class="hljs-built_in">random</span>.<span class="hljs-built_in">random</span>() * <span class="hljs-number">100</span>
<span class="hljs-comment"># Insert values to last column</span>
sheet.iloc[:<span class="hljs-number">21</span>, <span class="hljs-number">4</span>] = np.<span class="hljs-built_in">random</span>.<span class="hljs-built_in">random</span>(<span class="hljs-number">21</span>) * <span class="hljs-number">100</span>

<span class="hljs-comment"># === Insert computations ===</span>
<span class="hljs-comment"># Insert the computations on the row</span>
<span class="hljs-keyword">for</span> row_idx <span class="hljs-keyword">in</span> range(<span class="hljs-number">21</span>):
    <span class="hljs-comment"># I) Maximal value</span>
    sheet.iloc[row_idx, <span class="hljs-number">5</span>] = sheet.iloc[row_idx, <span class="hljs-number">0</span>:<span class="hljs-number">5</span>].<span class="hljs-built_in">max</span>()
    <span class="hljs-comment"># II) Minimal value</span>
    sheet.iloc[row_idx, <span class="hljs-number">6</span>] = sheet.iloc[row_idx, <span class="hljs-number">0</span>:<span class="hljs-number">5</span>].<span class="hljs-built_in">min</span>()
    <span class="hljs-comment"># III) Mean value</span>
    sheet.iloc[row_idx, <span class="hljs-number">7</span>] = sheet.iloc[row_idx, <span class="hljs-number">0</span>:<span class="hljs-number">5</span>].mean()
<span class="hljs-comment"># Insert the similar to rows:</span>
<span class="hljs-keyword">for</span> col_idx <span class="hljs-keyword">in</span> range(<span class="hljs-number">8</span>):
    <span class="hljs-comment"># I) Values of all</span>
    sheet.iloc[<span class="hljs-number">21</span>, col_idx] = sheet.iloc[<span class="hljs-number">0</span>:<span class="hljs-number">21</span>, col_idx].<span class="hljs-built_in">average</span>()
    <span class="hljs-comment"># II) Values of boys</span>
    sheet.iloc[<span class="hljs-number">22</span>, col_idx] = sheet.iloc[<span class="hljs-number">0</span>:<span class="hljs-number">11</span>, col_idx].<span class="hljs-built_in">average</span>()
    <span class="hljs-comment"># III) Values of girls</span>
    sheet.iloc[<span class="hljs-number">23</span>, col_idx] = sheet.iloc[<span class="hljs-number">11</span>:<span class="hljs-number">21</span>, col_idx].<span class="hljs-built_in">average</span>()

<span class="hljs-comment"># Export results to Excel file, <span class="hljs-doctag">TODO:</span> change the target directory:</span>
sheet.to_excel(<span class="hljs-string">"OUTPUTS/student_marks.xlsx"</span>, sheet_name=<span class="hljs-string">"Marks"</span>)

<span class="hljs-comment"># Top print table as Markdown</span>
print(sheet.to_markdown())
</code></pre>
"""

ENTITY = cr.Page(
    title="About Portable Spreadsheet",
    url_alias=None,  # Is homepage
    large_image_path=None,
    content=html_code,
    menu_position=0
)
