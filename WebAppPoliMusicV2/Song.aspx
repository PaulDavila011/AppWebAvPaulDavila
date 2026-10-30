<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV2.Song" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>PoliMusic - Songs</title>
    <link href="css/polimusic.css" rel="stylesheet" />
    <style>
        .song-layout { display: grid; grid-template-columns: minmax(280px, 380px) 1fr; gap: 24px; align-items: start; }
        .song-card { padding: 24px; }
        .song-card h2 { font-size: 17px; margin-bottom: 16px; }
        .song-audio { width: 230px; max-width: 100%; height: 34px; }
        .song-current { font-size: 12.5px; color: var(--pm-text-dim); word-break: break-all; }
        .song-topbar { display: flex; justify-content: space-between; align-items: center; gap: 12px; flex-wrap: wrap; margin-bottom: 24px; }
        @media (max-width: 900px) { .song-layout { grid-template-columns: 1fr; } }
    </style>
</head>
<body>
    <form id="formSong" runat="server">
        <main class="pm-content">
            <div class="song-topbar">
                <div>
                    <h1 class="pm-page-title">Songs</h1>
                    <p class="pm-page-sub">Upload, edit and listen to the PoliMusic catalog</p>
                </div>
                <div class="pm-topbar-actions">
                    <asp:Label ID="lblUser" runat="server" CssClass="pm-chip"></asp:Label>
                    <asp:HyperLink ID="lnkLogout" NavigateUrl="~/Logout.aspx" CssClass="pm-btn pm-btn-ghost pm-btn-sm" runat="server">Logout</asp:HyperLink>
                </div>
            </div>

            <asp:Panel ID="pnlMessage" Visible="false" runat="server">
                <asp:Label ID="lblMessage" runat="server"></asp:Label>
            </asp:Panel>

            <div class="song-layout">
                <section class="pm-card song-card">
                    <h2><asp:Literal ID="litFormTitle" runat="server" Text="New song"></asp:Literal></h2>
                    <asp:HiddenField ID="hfIdSong" Value="0" runat="server" />
                    <div class="pm-form">
                        <div class="pm-field">
                            <label class="pm-label" for="txtSongName">Song name <span class="pm-req">*</span></label>
                            <asp:TextBox ID="txtSongName" runat="server" CssClass="pm-input" MaxLength="50" placeholder="Song name"></asp:TextBox>
                        </div>
                        <div class="pm-field">
                            <label class="pm-label" for="fupSong">MP3 file <asp:Label ID="lblFileRequired" CssClass="pm-req" Text="*" runat="server"></asp:Label></label>
                            <asp:FileUpload ID="fupSong" runat="server" CssClass="pm-input" accept=".mp3,audio/mpeg" />
                            <span class="pm-hint">Only .mp3 files, max 50 MB.</span>
                            <asp:Label ID="lblCurrentFile" CssClass="song-current" runat="server"></asp:Label>
                        </div>
                        <div class="pm-form-actions">
                            <asp:Button ID="btnSave" runat="server" Text="Save" CssClass="pm-btn pm-btn-primary" OnClick="btnSave_Click" />
                            <asp:Button ID="btnCancel" runat="server" Text="Cancel" CssClass="pm-btn pm-btn-ghost" OnClick="btnCancel_Click" CausesValidation="false" />
                        </div>
                    </div>
                </section>

                <section>
                    <div class="pm-section-head">
                        <div>
                            <h2 class="pm-section-title">Catalog</h2>
                            <p class="pm-section-sub"><asp:Literal ID="litCount" runat="server"></asp:Literal></p>
                        </div>
                    </div>
                    <div class="pm-table-wrap">
                        <asp:GridView ID="gvSongs" runat="server" AutoGenerateColumns="false" DataKeyNames="ID_SONG"
                            CssClass="pm-table" GridLines="None" OnRowCommand="gvSongs_RowCommand"
                            EmptyDataText="There are no songs registered yet.">
                            <Columns>
                                <asp:BoundField DataField="ID_SONG" HeaderText="#" />
                                <asp:TemplateField HeaderText="Name">
                                    <ItemTemplate>
                                        <span class="pm-cell-main"><%# HttpUtility.HtmlEncode(Eval("SONG_NAME").ToString()) %></span>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Listen">
                                    <ItemTemplate>
                                        <audio class="song-audio" controls="controls" preload="none" src='<%# ResolveUrl(WebAppPoliMusicV2.Util.Media.ToMediaPath(Eval("SONG_PATH").ToString())) %>'></audio>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="PLAYS" HeaderText="Plays" />
                                <asp:TemplateField HeaderText="Actions">
                                    <ItemTemplate>
                                        <div class="pm-cell-actions">
                                            <asp:LinkButton ID="lnkEdit" runat="server" CommandName="EditSong" CommandArgument='<%# Eval("ID_SONG") %>'
                                                CssClass="pm-btn pm-btn-ghost pm-btn-sm" CausesValidation="false">Edit</asp:LinkButton>
                                            <asp:LinkButton ID="lnkDelete" runat="server" CommandName="DeleteSong" CommandArgument='<%# Eval("ID_SONG") %>'
                                                CssClass="pm-btn pm-btn-danger pm-btn-sm" CausesValidation="false"
                                                OnClientClick="return confirm('Delete this song?');">Delete</asp:LinkButton>
                                        </div>
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                        </asp:GridView>
                    </div>
                </section>
            </div>
        </main>
    </form>
</body>
</html>
