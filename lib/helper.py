import os
import xml.etree.cElementTree as etree
import re
import ctypes as ct
import glob

SHLoadIndirectString = ct.windll.shlwapi.SHLoadIndirectString
SHLoadIndirectString.argtypes = [ct.c_wchar_p, ct.c_wchar_p, ct.c_uint, ct.POINTER(ct.c_void_p)]
SHLoadIndirectString.restype = ct.HRESULT

RESOURCE_PREFIX = "ms-resource:"
RESOURCE_DISPLAY_FORMAT = "ms-resource://Windows.UI.SettingsAppThreshold/SystemSettings/Resources/{}/DisplayName"
RESOURCE_ALT_DISPLAY_FORMAT = "ms-resource://Windows.UI.SettingsAppThreshold" \
                              "/SystemSettings/Resources/{}/AlternateDisplayName"
RESOURCE_DESC_FORMAT = "ms-resource://Windows.UI.SettingsAppThreshold/SearchResources/{}/Description"
RESOURCE_SETTINGS_TITLE = "ms-resource://Windows.UI.SettingsAppThreshold/SystemSettings/Resources/SettingsAppTitle/Text"
RESOURCE_SETTINGS_TITLE2 = "ms-resource://Windows.UI.SettingsAppThreshold/resources/DisplayName"
RESOURCE_OPEN = "ms-resource://Windows.UI.ShellCommon/JumpViewUI/JumpView_CustomOpenAction"
RESOURCE_RUN_AS_ADMIN = "ms-resource://Windows.UI.ShellCommon/JumpViewUI/JumpView_CustomRunAsAdminAction"

WINDOWS10 = "http://schemas.microsoft.com/appx/manifest/foundation/windows10"
WINDOWS81 = "http://schemas.microsoft.com/appx/2013/manifest"
WINDOWS8 = "http://schemas.microsoft.com/appx/2010/manifest"


class AppXPackage(object):
    """Represents a windows app package
    """

    def __init__(self, property_dict):
        """Sets needed properties from the dict as member
        """
        # for key, value in property_dict.items():
        #     setattr(self, key, value)

        self.Name = property_dict["Name"] if "Name" in property_dict else None
        self.InstallLocation = property_dict["InstallLocation"] if "InstallLocation" in property_dict else None
        self.PackageFamilyName = property_dict["PackageFamilyName"] if "PackageFamilyName" in property_dict else None
        self.applications = []

    def apps(self):
        if not self.applications:
            self.applications = self._get_applications()
        return self.applications

    def _get_applications(self):
        """Reads the manifest of the package and extracts name, description, applications and logos
        """
        if not self.InstallLocation:
            return []

        manifest_path = os.path.join(self.InstallLocation, "AppxManifest.xml")
        if not os.path.isfile(manifest_path):
            return []
        manifest = etree.parse(manifest_path)
        ns = {"default": re.sub(r"{(.*?)}.+", r"\1", manifest.getroot().tag)}

        package_applications = manifest.findall("./default:Applications/default:Application", ns)
        if not package_applications:
            return []

        apps = []

        package_description = ""
        default_description_node = manifest.find("./default:Properties/default:Description", ns)
        if default_description_node is not None:
            package_description = default_description_node.text.strip()

        package_display_name = ""
        default_display_name_node = manifest.find("./default:Properties/default:DisplayName", ns)
        if default_display_name_node is not None:
            package_display_name = default_display_name_node.text.strip()

        package_icon_path = ""
        logo_node = manifest.find("./default:Properties/default:Logo", ns)
        if logo_node is not None:
            logo = logo_node.text
            package_icon_path = os.path.join(self.InstallLocation, logo)

        for application in package_applications:
            app_display_name = ""
            app_description = ""
            app_icon_path = ""
            app_misc = False

            visual_elements = next((elem for elem in application if elem.tag.endswith("VisualElements")), None)
            if visual_elements is not None:
                default_tile = next((elem for elem in visual_elements if elem.tag.endswith('DefaultTile')), None)

                app_misc = visual_elements.get("AppListEntry") == "none" \
                    if "AppListEntry" in visual_elements.attrib else False

                app_display_name = visual_elements.get("DisplayName")
                app_description = visual_elements.get("Description")

                logos = {attr: visual_elements.get(attr) for attr in visual_elements.attrib if "logo" in attr.lower()}
                if ns["default"] == WINDOWS10 and "Square44x44Logo" in logos:
                    app_icon_path = os.path.join(self.InstallLocation, logos["Square44x44Logo"])
                elif ns["default"] == WINDOWS81 and "Square30x30Logo" in logos:
                    app_icon_path = os.path.join(self.InstallLocation, logos["Square30x30Logo"])
                elif ns["default"] == WINDOWS8 and "SmallLogo" in logos:
                    app_icon_path = os.path.join(self.InstallLocation, logos["SmallLogo"])
                else:
                    if default_tile is not None:
                        logos.update({attr: default_tile.get(attr) for attr in default_tile.attrib
                                      if "logo" in attr.lower()})
                    square_logos = {key: value for key, value in logos.items() if "square" in key.lower()}
                    wide_logos = {key: value for key, value in logos.items() if "wide" in key.lower()}

                    if square_logos:
                        biggest = max(square_logos.keys(), key=lambda x: int(re.search(r"(\d+)x\d+", x).groups()[0]))
                        app_icon_path = os.path.join(self.InstallLocation, logos[biggest])
                    elif not app_icon_path and wide_logos:
                        biggest = max(wide_logos, key=lambda x: re.search(r"(\d+)x\d+", x).groups()[0])
                        app_icon_path = os.path.join(self.InstallLocation, logos[biggest])
                    elif not app_icon_path and logos:
                        biggest = min(logos)
                        app_icon_path = os.path.join(self.InstallLocation, logos[biggest])
                    elif not app_icon_path:
                        app_icon_path = package_icon_path

                if app_display_name and app_display_name.startswith(RESOURCE_PREFIX):
                    resource = self.get_resource(self.InstallLocation, app_display_name, self.Name)
                    if resource:
                        app_display_name = resource
                    else:
                        if package_display_name and package_display_name.startswith(RESOURCE_PREFIX):
                            resource = self.get_resource(self.InstallLocation, package_display_name, self.Name)
                            if resource:
                                package_display_name = resource
                                app_display_name = package_display_name
                            else:
                                app_display_name = self.Name
                        else:
                            app_display_name = self.Name

                if app_description and app_description.startswith(RESOURCE_PREFIX):
                    resource = self.get_resource(self.InstallLocation, app_description, self.Name)
                    if resource:
                        app_description = resource
                    else:
                        if package_description and package_description.startswith(RESOURCE_PREFIX):
                            resource = self.get_resource(self.InstallLocation, package_description, self.Name)
                            if resource:
                                package_description = resource
                                app_description = package_description

                apps.append(AppX(execution="shell:AppsFolder\\{}!{}".format(self.PackageFamilyName, application.get("Id")),
                                 display_name=app_display_name,
                                 description=app_description,
                                 icon_path=app_icon_path,
                                 app_id="{}!{}".format(self.PackageFamilyName, application.get("Id")),
                                 misc_app=app_misc))
        return apps

    @staticmethod
    def get_resource(install_location, resource, name=None):
        """Helper method to resolve resource strings to their (localized) value
        """
        pri_files = []
        pri_files.extend(glob.glob(install_location + '/*.pri'))
        pri_files.extend(glob.glob(install_location + '/**/*.pri'))
        # the assumption is, that localized .pri resource data files are deeper in the file tree
        # and therefore have a longer path
        pri_files_sorted = sorted(pri_files, key=lambda file: len(file), reverse=True)
        resource_root_names = ["/resources"]
        if name:
            resource_root_names.append(name)

        for pri_file in pri_files_sorted:
            for resource_root_name in resource_root_names:
                try:
                    if resource[0:12] == RESOURCE_PREFIX:
                        resource_key = resource[12:]
                        if resource_key.startswith("//"):
                            resource_path = resource
                        elif resource_key.startswith("/"):
                            resource_path = RESOURCE_PREFIX + "//" + resource_key
                        else:
                            resource_path = RESOURCE_PREFIX + "//" + resource_root_name + "/" + resource_key

                        resource_descriptor = "@{{{}? {}}}".format(pri_file, resource_path)

                        inp = ct.create_unicode_buffer(resource_descriptor)
                        output = ct.create_unicode_buffer(1024)
                        result = SHLoadIndirectString(inp, output, ct.sizeof(output), None)
                        if result == 0 and output.value:
                            if not output.value.startswith(RESOURCE_PREFIX):
                                return output.value
                except OSError:
                    pass

        return None


class AppX(object):
    """Represents an executable application from a windows app package
    """

    def __init__(self, execution=None, display_name=None, description=None, icon_path=None, app_id=None,
                 misc_app=False):
        self.execution = execution
        self.display_name = display_name
        self.description = description
        self.icon_path = icon_path
        self.app_id = app_id
        self.misc_app = misc_app


if __name__ == "__main__":
    import json
    import subprocess

    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    output, err = subprocess.Popen(["powershell.exe",
                                    "-noprofile",
                                    "chcp 65001 >$null; [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new(); Get-AppxPackage | ConvertTo-Json"],
                                    stdout=subprocess.PIPE,
                                    universal_newlines=False,
                                    shell=False,
                                    startupinfo=startupinfo).communicate()

    catalog = []
    packages = json.loads(output.decode("utf8", "replace"))
    for package in packages:
        p = AppXPackage(package)
        apps = p.apps()
        if all([app.misc_app for app in apps]):
            continue
        print(f"{p.Name:50} {p.InstallLocation}")
        for app in apps:
            if not app.misc_app:
                print(f"  - {app.display_name} ({app.description})")
