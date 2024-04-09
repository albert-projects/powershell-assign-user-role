---


---

<h1 id="active-directory-user-management-gui-script">Active Directory User Management GUI Script</h1>
<p>This project presents a PowerShell script with a graphical user interface (GUI) designed to facilitate the management of domain users, roles, and Organizational Units (OUs) within a Windows Active Directory environment. Specifically aimed at assisting HR staff in assigning access rights to new casual staff members, the script offers functionality to set both the “Weekday” and “End date” for the access privileges.</p>
<h2 id="key-features">Key Features</h2>
<ul>
<li>
<p><strong>Graphical Interface:</strong> The script provides a user-friendly GUI, enabling HR staff to update user information, roles, and OUs efficiently.</p>
</li>
<li>
<p><strong>Access Control:</strong> HR personnel can easily assign access rights to new casual staff, with the ability to specify start and end dates for their access.</p>
</li>
<li>
<p><strong>Active Directory Integration:</strong> The script requires the machine to have the Active Directory Module installed and be connected to an Active Directory environment for seamless operation.</p>
</li>
</ul>
<h2 id="usage">Usage</h2>
<ol>
<li>Ensure that the machine running the script has the Active Directory Module installed.</li>
<li>Clone or download the script to a local directory.</li>
<li>Run the script on a machine connected to the Active Directory environment.</li>
<li>Utilize the GUI interface to update user details, roles, and OUs as required.</li>
<li>Set access rights for new casual staff by specifying the desired start and end dates.</li>
</ol>
<h2 id="prerequisites">Prerequisites</h2>
<ul>
<li><strong>Active Directory Module:</strong> Ensure that the machine has the Active Directory Module installed.</li>
<li><strong>Windows Environment:</strong> The script is intended for use in a Windows environment, connected to an Active Directory domain.</li>
</ul>
<h2 id="getting-started">Getting Started</h2>
<ol>
<li>Clone or download the repository to your local machine.</li>
<li>Open PowerShell with administrative privileges.</li>
<li>Navigate to the directory containing the script.</li>
<li>Execute the script using the following command:</li>
</ol>
<pre class=" language-powershell"><code class="prism  language-powershell">    <span class="token punctuation">.</span><span class="token operator">/</span>AssignUserRole<span class="token punctuation">.</span>ps1
</code></pre>
<ol start="5">
<li>Follow the prompts in the GUI to manage domain users, roles, and OUs effectively.</li>
</ol>
<h2 id="section"><img src="https://github.com/albert-projects/powershell-assign-user-role/blob/master/screenshot.png" alt="Screenshot"></h2>
<p>powershell-assign-user-role<br>
Albert Kwan</p>

