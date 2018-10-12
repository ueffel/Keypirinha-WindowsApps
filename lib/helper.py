import os
import xml.etree.ElementTree as etree
import re
import ctypes as ct

SHLoadIndirectString = ct.windll.shlwapi.SHLoadIndirectString
SHLoadIndirectString.argtypes = [ct.c_wchar_p, ct.c_wchar_p, ct.c_uint, ct.POINTER(ct.c_void_p)]
SHLoadIndirectString.restype = ct.HRESULT


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

    async def apps(self):
        if not self.applications:
            self.applications = await self._get_applications()
        return self.applications

    async def _get_applications(self):
        """Reads the manifest of the package and extracts name, description, applications and logos
        """
        manifest_path = os.path.join(self.InstallLocation, "AppxManifest.xml")
        if not os.path.isfile(manifest_path):
            return []
        manifest = etree.parse(manifest_path)
        ns = {"default": re.sub(r"\{(.*?)\}.+", r"\1", manifest.getroot().tag)}

        package_applications = manifest.findall("./default:Applications/default:Application", ns)
        if not package_applications:
            return []

        apps = []

        package_identity = None
        package_identity_node = manifest.find("./default:Identity", ns)
        if package_identity_node is not None:
            package_identity = package_identity_node.get("Name")

        description = None
        description_node = manifest.find("./default:Properties/default:Description", ns)
        if description_node is not None:
            description = description_node.text

        display_name = None
        display_name_node = manifest.find("./default:Properties/default:DisplayName", ns)
        if display_name_node is not None:
            display_name = display_name_node.text

        icon_path = None
        logo_node = manifest.find("./default:Properties/default:Logo", ns)
        if logo_node is not None:
            logo = logo_node.text
            icon_path = os.path.join(self.InstallLocation, logo)

        for application in package_applications:
            if display_name and display_name.startswith("ms-resource:"):
                resource = self._get_resource(self.InstallLocation, package_identity, display_name)
                if resource is not None:
                    display_name = resource
                else:
                    continue

            if description and description.startswith("ms-resource:"):
                resource = self._get_resource(self.InstallLocation, package_identity, description)
                if resource is not None:
                    description = resource
                else:
                    continue

            apps.append(AppX("shell:AppsFolder\{}!{}".format(self.PackageFamilyName, application.get("Id")),
                             display_name,
                             description,
                             icon_path))
        return apps

    @staticmethod
    def _get_resource(install_location, package_id, resource):
        """Helper method to resolve resource strings to their (localized) value
        """
        try:
            resource_descriptor = None
            if resource.startswith("ms-resource:/"):
                resource_descriptor = "@{{{}\\resources.pri? {}}}".format(install_location,
                                                                          resource)
            elif resource.startswith("ms-resource:"):
                resource_descriptor = "@{{{}\\resources.pri? ms-resource://{}/resources/{}}}".format(install_location,
                                                                                                     package_id,
                                                                                                     resource[len("ms-resource:"):])
            if resource_descriptor is None:
                return None

            inp = ct.create_unicode_buffer(resource_descriptor)
            output = ct.create_unicode_buffer(1024)
            result = SHLoadIndirectString(inp, output, ct.sizeof(output), None)
            if result == 0:
                return output.value
            else:
                return None
        except OSError:
            pass

        try:
            resource_descriptor = "@{{{}\\resources.pri? ms-resource://{}}}".format(install_location,
                                                                                    resource[len("ms-resource:"):])
            input = ct.create_unicode_buffer(resource_descriptor)
            output = ct.create_unicode_buffer(1024)
            result = SHLoadIndirectString(inp, output, ct.sizeof(output), None)
            if result == 0:
                return output.value
            else:
                return None
        except OSError:
            pass

        return None


class AppX(object):
    """Represents an executable application from a windows app package
    """
    def __init__(self, execution=None, display_name=None, description=None, icon_path=None):
        self.execution = execution
        self.display_name = display_name
        self.description = description
        self.icon_path = icon_path
