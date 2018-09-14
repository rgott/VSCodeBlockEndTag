// Thanks to https://github.com/jaredpar/ControlCharAdornmentSample/blob/master/CharDisplayTaggerSource.cs

using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Text.Tagging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Collections.ObjectModel;
using Microsoft.VisualStudio.Editor;

namespace CodeBlockEndTag
{

    /// <summary> 
    /// This tagger provides editor tags that are inserted into the TextView (IntraTextAdornmentTags)
    /// The tags are added after each code block encapsulated by curly bracets: { ... }
    /// The tags will show the code blocks condition, or whatever serves as header for the block
    /// By clicking on a tag, the editor will jump to that code blocks header
    /// </summary>
    internal class CBETagger : ITagger<IntraTextAdornmentTag>, IDisposable
    {
        private static readonly ReadOnlyCollection<ITagSpan<IntraTextAdornmentTag>> EmptyTagCollection =
            new ReadOnlyCollection<ITagSpan<IntraTextAdornmentTag>>(new List<ITagSpan<IntraTextAdornmentTag>>());

        #region Properties & Fields

        // EventHandler for ITagger<IntraTextAdornmentTag> tags changed event
        EventHandler<SnapshotSpanEventArgs> _changedEvent;

        /// <summary>
        /// Service by VisualStudio for fast searches in texts 
        /// </summary>
        readonly ITextSearchService _TextSearchService;

        /// <summary>
        /// Service by VisualStudio for fast navigation in structured texts
        /// </summary>
        readonly ITextStructureNavigator _TextStructureNavigator;

        /// <summary>
        /// The TextView this tagger is assigned to
        /// </summary>
        readonly IWpfTextView _TextView;

        /// <summary>
        /// This is a list of already created adornment tags used as cache
        /// </summary>
        readonly List<CBAdornmentData> _adornmentCache = new List<CBAdornmentData>();


        /// <summary>
        /// Used to get the editor font size
        /// </summary>
        readonly IVsFontsAndColorsInformation _VSFontsInformation;

        /// <summary>
        /// FontSize used by tags
        /// </summary>
        double _FontSize = 9;

        /// <summary>
        /// This is the visible span of the textview
        /// </summary>
        Span? _VisibleSpan;

        /// <summary>
        /// Is set, when the instance is disposed
        /// </summary>
        bool _Disposed { get; set; }

        #endregion

        #region Ctor

        /// <summary>
        /// Creates a new instance of CBRTagger
        /// </summary>
        /// <param name="provider">the CBETaggerProvider that created the tagger</param>
        /// <param name="textView">the WpfTextView this tagger is assigned to</param>
        /// <param name="sourceBuffer">the TextBuffer this tagger should work with</param>
        internal CBETagger(CBETaggerProvider provider, IWpfTextView textView)
        {
            if (provider == null || textView == null)
                throw new ArgumentNullException("The arguments of CBETagger can't be null");

            _TextView = textView;

            // Getting services provided by VisualStudio
            _TextStructureNavigator = provider.GetTextStructureNavigator(_TextView.TextBuffer);
            _TextSearchService = provider.TextSearchService;
            _VSFontsInformation = TryGetFontAndColorInfo(provider.VsFontsAndColorsInformationService);

            // Hook up events
            _TextView.TextBuffer.Changed += TextBuffer_Changed;
            _TextView.LayoutChanged += OnTextViewLayoutChanged;
            _TextView.Caret.PositionChanged += Caret_PositionChanged;
            CBETagPackage.Instance.PackageOptionChanged += OnPackageOptionChanged;
            if (_VSFontsInformation != null)
            {
                ReloadFontSize(_VSFontsInformation);
                _VSFontsInformation.Updated += _VSFontsInformation_Updated;
            }
        }

        #endregion

        #region VsFontAndColorInformation

        private static IVsFontsAndColorsInformation TryGetFontAndColorInfo(IVsFontsAndColorsInformationService service)
        {
            var guidTextFileType = new Guid(2184822468u, 61063, 4560, 140, 152, 0, 192, 79, 194, 171, 34);
            var fonts = new FontsAndColorsCategory(
                guidTextFileType,
                DefGuidList.guidTextEditorFontCategory,
                DefGuidList.guidTextEditorFontCategory);
            return service?.GetFontAndColorInformation(fonts);
        }

        private void ReloadFontSize(IVsFontsAndColorsInformation _VSFontsInformation)
        {
            var pref = _VSFontsInformation.GetFontAndColorPreferences();
            var font = System.Drawing.Font.FromHfont(pref.hRegularViewFont);
            _FontSize = font?.Size ?? _FontSize;
        }

        private void _VSFontsInformation_Updated(object sender, EventArgs e)
        {
            if (_VSFontsInformation != null)
            {
                ReloadFontSize(_VSFontsInformation);
            }
        }

        #endregion

        #region TextBuffer changed

        private void Caret_PositionChanged(object sender, CaretPositionChangedEventArgs e)
        {
            var start = Math.Min(e.OldPosition.BufferPosition.Position, e.NewPosition.BufferPosition.Position);
            var end = Math.Max(e.OldPosition.BufferPosition.Position, e.NewPosition.BufferPosition.Position);
            if (start != end)
            {
                InvalidateSpan(new Span(start, end - start));
            }
        }

        void TextBuffer_Changed(object sender, TextContentChangedEventArgs e)
        {
            foreach (var textChange in e.Changes)
            {
                OnTextChanged(textChange);
            }
        }

        void OnTextChanged(ITextChange textChange)
        {
            // remove or update tags in adornment cache
            var remove = new List<CBAdornmentData>();
            foreach (var adornment in _adornmentCache)
            {
                if (!(adornment.HeaderStartPosition > textChange.OldEnd || adornment.EndPosition < textChange.OldPosition))
                {
                    remove.Add(adornment);
                }
                else if (adornment.HeaderStartPosition > textChange.OldEnd)
                {
                    adornment.HeaderStartPosition += textChange.Delta;
                    adornment.StartPosition += textChange.Delta;
                    adornment.EndPosition += textChange.Delta;
                }
            }

            foreach (var adornment in remove)
            {
                RemoveFromCache(adornment);
            }
        }

        private void RemoveFromCache(CBAdornmentData adornment)
        {
            if (adornment.Adornment is CBETagControl tag)
            {
                tag.TagClicked -= Adornment_TagClicked;
            }
            _adornmentCache.Remove(adornment);
        }

        #endregion

        #region ITagger<IntraTextAdornmentTag>

        IEnumerable<ITagSpan<IntraTextAdornmentTag>> ITagger<IntraTextAdornmentTag>.GetTags(NormalizedSnapshotSpanCollection spans)
        {
            return spans.SelectMany(GetTags);
        }

        event EventHandler<SnapshotSpanEventArgs> ITagger<IntraTextAdornmentTag>.TagsChanged
        {
            add { _changedEvent += value; }
            remove { _changedEvent -= value; }
        }

        #endregion

        #region Tag placement

        internal ReadOnlyCollection<ITagSpan<IntraTextAdornmentTag>> GetTags(SnapshotSpan span)
        {
            if (!CBETagPackage.CBETaggerEnabled 
                || span.Snapshot != _TextView.TextBuffer.CurrentSnapshot 
                || span.Length == 0)
            {
                return EmptyTagCollection;
            }

            // if big span, return only tags for visible area
            if (span.Length > 1000 && _VisibleSpan.HasValue)
            {
                var overlap = span.Overlap(_VisibleSpan.Value);
                if (overlap.HasValue)
                {
                    span = overlap.Value;
                    if (span.Length == 0)
                        return EmptyTagCollection;
                }
            }
#if DEBUG
            return Measure(() => GetTagsCore(span), (time) =>
            {
                return "Time elapsed: " + time +
                    " on Thread: " + System.Threading.Thread.CurrentThread.ManagedThreadId +
                    " in Span: " + span.Start.Position + ":" + span.End.Position + " length: " + span.Length;
            });

#else
            return GetTagsCore(span);
#endif
        }

        private static T Measure<T>(Func<T> func, Func<TimeSpan, string> output)
        {
            var watch = new System.Diagnostics.Stopwatch();
            watch.Start();
            var result = func.Invoke();
            watch.Stop();
            if (watch.Elapsed.Milliseconds > 100)
            {
                System.Diagnostics.Debug.WriteLine(output.Invoke(watch.Elapsed));
            }
            return result;
        }

        private ReadOnlyCollection<ITagSpan<IntraTextAdornmentTag>> GetTagsCore(SnapshotSpan span)
        {
            var list = new List<ITagSpan<IntraTextAdornmentTag>>();
            var offset = span.Start.Position;
            var snapshot = span.Snapshot;

            // vars used in loop
            
            var isSingleLineComment = false;
            var isMultiLineComment = false;

            // Find all closing bracets
            for (int i = 0; i < span.Length; i++)
            {
                var position = i + offset;

                // Skip comments
                switch (snapshot[position])
                {
                    case '/':
                        if (position > 0)
                        {
                            if (snapshot[position - 1] == '/')
                                isSingleLineComment = true;
                            if (snapshot[position - 1] == '*')
                            {
                                if (!isMultiLineComment)
                                {
                                    // Multiline comment was not started in this span
                                    // Every tag until now was inside a comment
                                    foreach (var tag in list)
                                    {
                                        RemoveFromCache((tag.Tag.Adornment as CBETagControl).AdornmentData);
                                    }
                                    list.Clear();
                                }
                                isMultiLineComment = false;
                            }
                        }
                        break;
                    case '*':
                        if (position > 0 && snapshot[position - 1] == '/')
                            isMultiLineComment = true;
                        break;
                    case '\n':
                    case '\r':
                        isSingleLineComment = false;
                        break;
                }

                if (snapshot[position] != '}' || isSingleLineComment || isMultiLineComment)
                    continue;


                SnapshotSpan cbSpan;
                int cbStartPosition;
                // getting start and end position of code block
                var cbEndPosition = position;
                if (position >= 0 && snapshot[position - 1] == '{')
                {
                    // empty code block {} 
                    cbStartPosition = position - 1;
                    cbSpan = new SnapshotSpan(snapshot, position - 1, 1);
                }
                else
                {
                    // create inner span to navigate to get code block start
                    cbSpan = _TextStructureNavigator.GetSpanOfEnclosing(new SnapshotSpan(snapshot, position - 1, 1));
                    cbStartPosition = cbSpan.Start;
                }

                // Don't display tag for code blocks on same line
                if (!snapshot.GetText(cbSpan).Contains('\n'))
                    continue;

                // getting the code blocks header 
                var cbHeaderPosition = -1;
                string cbHeader;

                if (snapshot[cbStartPosition] == '{')
                {
                    // cbSpan does not contain the header
                    cbHeader = GetCodeBlockHeader(cbSpan, out cbHeaderPosition);
                }
                else
                {
                    // cbSpan does contain the header
                    cbHeader = GetCodeBlockHeader(cbSpan, out cbHeaderPosition, position);
                }

                // Trim header
                if (cbHeader != null && cbHeader.Length > 0)
                {
                    cbHeader = cbHeader.Trim()
                        .Replace(Environment.NewLine, "")
                        .Replace('\t', ' ');
                    // Strip unnecessary spaces
                    while (cbHeader.Contains("  "))
                    {
                        cbHeader = cbHeader.Replace("  ", " ");
                    }
                }

                // Skip tag if option "only when header not visible"
                if (_VisibleSpan != null && !IsTagVisible(cbHeaderPosition, cbEndPosition, _VisibleSpan, snapshot))
                    continue;

                var iconMoniker = Microsoft.VisualStudio.Imaging.KnownMonikers.QuestionMark;
                if (CBETagPackage.CBEDisplayMode != (int)CBEOptionPage.DisplayModes.Text &&
                    !string.IsNullOrWhiteSpace(cbHeader) && !cbHeader.Contains("{"))
                {
                    iconMoniker = IconMonikerSelector.SelectMoniker(cbHeader);
                }

                // use cache or create new tag
                var cbAdornmentData = _adornmentCache
                    .Find(x =>
                    x.StartPosition == cbStartPosition
                    && x.EndPosition == cbEndPosition);

                CBETagControl tagElement;
                if (cbAdornmentData?.Adornment != null)
                {
                    tagElement = cbAdornmentData.Adornment as CBETagControl;
                }
                else
                {
                    // create new adornment
                    tagElement = new CBETagControl()
                    {
                        Text = cbHeader,
                        IconMoniker = iconMoniker,
                        DisplayMode = CBETagPackage.CBEDisplayMode
                    };
                    
                    tagElement.TagClicked += Adornment_TagClicked;
                    
                    cbAdornmentData = new CBAdornmentData(cbStartPosition, cbEndPosition, cbHeaderPosition, tagElement);
                    tagElement.AdornmentData = cbAdornmentData;
                    _adornmentCache.Add(cbAdornmentData);
                }

                tagElement.LineHeight = _FontSize * CBETagPackage.CBETagScale;

                // Add new tag to list
                var cbTag = new IntraTextAdornmentTag(tagElement, null);
                var cbSnapshotSpan = new SnapshotSpan(snapshot, position + 1, 0);
                var cbTagSpan = new TagSpan<IntraTextAdornmentTag>(cbSnapshotSpan, cbTag);
                list.Add(cbTagSpan);
            }

            return new ReadOnlyCollection<ITagSpan<IntraTextAdornmentTag>>(list);
        }

        /// <summary>
        /// Capture the header of a code block
        /// Returns the text and outputs the start position within the snapshot
        /// </summary>
        string GetCodeBlockHeader(SnapshotSpan cbSpan, out int headerStart, int maxEndPosition = 0)
        {
            if (maxEndPosition == 0)
                maxEndPosition = cbSpan.Start;
            var snapshot = cbSpan.Snapshot;
            var currentSpan = cbSpan;

            // set end of header to first start of code block {
            for (int i = cbSpan.Start; i < cbSpan.End; i++)
            {
                if (snapshot[i] == '{')
                {
                    maxEndPosition = i;
                    break;
                }
            }

            Span headerSpan, headerSpan2;
            string headerText, headerText2;
            var loops = 0;
            // check all enclosing spans until the header is complete
            do
            {
                // get text of current span
                headerStart = currentSpan.Start;
                headerSpan = new Span(headerStart, Math.Min(maxEndPosition, currentSpan.Span.End) - headerStart);
                if (headerSpan.Length == 0)
                    continue;
                headerText = snapshot.GetText(headerSpan);

                // found header if it begins with a letter or contains a lambda
                if (!string.IsNullOrWhiteSpace(headerText))
                //&& (char.IsLetter(headerText[0]) || headerText[0]=='[' || headerText.Contains("=>")))
                {
                    // recognize "else if" too
                    if (headerText.StartsWith("if") && ((currentSpan = _TextStructureNavigator.GetSpanOfEnclosing(currentSpan)) != null))
                    {
                        // check what comes before the "if"
                        headerSpan2 = new Span(currentSpan.Start, Math.Min(maxEndPosition, currentSpan.Span.End) - currentSpan.Start);
                        headerText2 = snapshot.GetText(headerSpan2);
                        if (headerText2.StartsWith("else"))
                        {
                            headerStart = headerSpan2.Start;
                            headerText = headerText2;
                        }
                    }
                    else if (headerText.Contains('\r') || headerText.Contains('\n'))
                    {
                        // skip annotations
                        headerText = headerText.Replace('\r', '\n').Replace("\n\n", "\n");
                        var headerLines = headerText.Split('\n');
                        var annotaions = true;
                        var openBracets = 0;
                        headerText = string.Empty;
                        foreach (var line in headerLines)
                        {
                            if (!string.IsNullOrWhiteSpace(line))
                            {
                                var trimmedline = line.Trim();
                                if (annotaions && (trimmedline[0] == '[' || openBracets > 0))
                                {
                                    openBracets += trimmedline.Count(c => c == '[');
                                    openBracets -= trimmedline.Count(c => c == ']');
                                    continue;
                                }
                                annotaions = false;
                                if (!string.IsNullOrWhiteSpace(headerText))
                                    headerText += Environment.NewLine;
                                headerText += trimmedline;
                            }
                        }
                    }
                    return headerText;
                }
                currentSpan = _TextStructureNavigator.GetSpanOfEnclosing(currentSpan);
            } while (loops++ > 10); // TODO: create better algorithm than looping 10 times

            // No header found
            headerStart = -1;
            return null;
        }


#endregion

#region Tag Clicked Handler

        /// <summary>
        /// Handles the click event on a tag
        /// </summary>
        private void Adornment_TagClicked(CBAdornmentData adornment, bool jumpToHead)
        {
            if (_TextView == null)
            {
                return;
            }

            SnapshotPoint targetPoint;
            if (jumpToHead)
            {
                // Jump to header
                targetPoint = new SnapshotPoint(_TextView.TextBuffer.CurrentSnapshot, adornment.HeaderStartPosition);
                _TextView.DisplayTextLineContainingBufferPosition(targetPoint, 30, ViewRelativePosition.Top);
            }
            else
            {
                // Set caret behind closing bracet
                targetPoint = new SnapshotPoint(_TextView.TextBuffer.CurrentSnapshot, adornment.EndPosition + 1);
            }
            _TextView.Caret.MoveTo(targetPoint);
        }

#endregion

#region Options changed

        /// <summary>
        /// Handles the event when any package option is changed
        /// </summary>
        private void OnPackageOptionChanged(object sender)
        {
            var start = Math.Max(0, _VisibleSpan.HasValue ? _VisibleSpan.Value.Start : 0);
            var end = Math.Max(1, _VisibleSpan.HasValue ? _VisibleSpan.Value.End : 1);

            var span = new Span(start, end - start);
            ClearCache(span);
            InvalidateSpan(span);
        }

        private void ClearCache(Span invalidateSpan)
        {
            _adornmentCache
                    .Where(a => a.HeaderStartPosition >= invalidateSpan.Start || a.EndPosition >= invalidateSpan.Start)
                    .ToList().ForEach(a => RemoveFromCache(a));
        }

        /// <summary>
        /// Invalidates all cached tags within or after the given span
        /// </summary>
        private void InvalidateSpan(Span invalidateSpan)
        {
            // Invalidate span
            if (invalidateSpan.End <= _TextView.TextBuffer.CurrentSnapshot.Length)
            {
                _changedEvent?.Invoke(this, new SnapshotSpanEventArgs(
                    new SnapshotSpan(_TextView.TextBuffer.CurrentSnapshot, invalidateSpan)));
            }
        }

#endregion

#region IDisposable

        private void Dispose(bool disposing)
        {
            if (_Disposed)
                return;

            if (disposing)
            {
                // Clean up all events and references
                CBETagPackage.Instance.PackageOptionChanged -= OnPackageOptionChanged;
                _TextView.LayoutChanged -= OnTextViewLayoutChanged;
                _TextView.TextBuffer.Changed -= TextBuffer_Changed;
            }
            _Disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

#endregion

#region visibility of tags

        /// <summary>
        /// Checks if a tag's header is visible
        /// </summary>
        /// <param name="start">Start position of code block</param>
        /// <param name="end">End position of code block</param>
        /// <param name="visibleSpan">the visible span in the textview</param>
        /// <param name="snapshot">reference to text snapshot. Used for caret check</param>
        /// <returns>true if the tag is visible (or if all tags are shown)</returns>
        private bool IsTagVisible(int start, int end, Span? visibleSpan, ITextSnapshot snapshot)
        {
            var isVisible = false;
            // Check general condition
            if (CBETagPackage.CBEVisibilityMode == (int)CBEOptionPage.VisibilityModes.Always
                || !visibleSpan.HasValue)
            {
                isVisible = true;
            }
            // Check visible span
            if (!isVisible)
            {
                var val = visibleSpan.Value;
                isVisible = (start < val.Start && end >= val.Start && end <= val.End);
            }
            // Check if caret is in this line
            if (isVisible && _TextView != null)
            {
                var caretIndex = _TextView.Caret.Position.BufferPosition.Position;
                var lineStart = Math.Min(caretIndex, end);
                var lineEnd = Math.Max(caretIndex, end);
                if (lineStart == lineEnd)
                {
                    return false;
                }
                else if (lineStart >= 0 && lineEnd <= snapshot.Length)
                {
                    var line = snapshot.GetText(lineStart, lineEnd - lineStart);
                    return line.Contains('\n');
                }
            }
            return isVisible;
        }

        /// <summary>
        /// Returns the visible span for the given textview
        /// </summary>
        private Span? GetVisibleSpan(ITextView textView)
        {
            if (textView?.TextViewLines != null && textView.TextViewLines.Count > 2)
            {
                // Index 0 not yet visible
                var firstVisibleLine = textView.TextViewLines[1];
                // Last index not visible, too
                var lastVisibleLine = textView.TextViewLines[textView.TextViewLines.Count - 2];

                return new Span(firstVisibleLine.Start, lastVisibleLine.End - firstVisibleLine.Start);
            }
            return null;
        }
#endregion

#region TextView scrolling

        private void OnTextViewLayoutChanged(object sender, TextViewLayoutChangedEventArgs e)
        {
            // get new visible span
            var visibleSpan = GetVisibleSpan(_TextView);
            if (!visibleSpan.HasValue)
                return;

            // only if new visible span is different from old
            if (!_VisibleSpan.HasValue
                || _VisibleSpan.Value.Start != visibleSpan.Value.Start
                || _VisibleSpan.Value.End < visibleSpan.Value.End)
            {
                // invalidate new and/or old visible span
                var invalidSpans = new List<Span>();
                var newSpan = visibleSpan.Value;
                if (!_VisibleSpan.HasValue)
                {
                    invalidSpans.Add(newSpan);
                }
                else
                {
                    var oldSpan = _VisibleSpan.Value;
                    // invalidate two spans if old and new do not overlap
                    if (newSpan.Start > oldSpan.End || newSpan.End < oldSpan.Start)
                    {
                        invalidSpans.Add(newSpan);
                        invalidSpans.Add(oldSpan);
                    }
                    else
                    {
                        // invalidate one big span (old and new joined)
                        invalidSpans.Add(newSpan.Join(oldSpan));
                    }
                }

                _VisibleSpan = visibleSpan;

                // refresh tags
                foreach (var span in invalidSpans)
                {
                    if (CBETagPackage.CBEVisibilityMode != (int)CBEOptionPage.VisibilityModes.Always)
                    {
                        InvalidateSpan(span);
                    }
                }
            }
        }
#endregion
    }
}
