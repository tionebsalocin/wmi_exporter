// returns data points from Test_Class

// Custom WMI Class (Test_Class class)

package collector

import (
	"flag"
	"fmt"
	"log"
	"strings"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/prometheus/client_golang/prometheus"
)

func init() {
	Factories["custom"] = NewCustomCollector
}

var (
	customClassPath = flag.String("collector.custom.classpath", "", "List WMI Class Path: root\\cimv2:win32_service;...")
)

type wmiClass struct {
	Name      string
	Instances []wmiInstance
}

func (c *wmiClass) Export() {
	for _, v := range c.Instances {
		for key, value := range v.Properties {
			fmt.Println("Class:", c.Name, "Instance:", v.Name, "Propertie:", key, "Value:", value)
		}
	}
}

type wmiInstance struct {
	Name       string
	Properties map[string]interface{}
}

func (i *wmiInstance) Export() {
	for key, value := range i.Properties {
		fmt.Println("Instance:", i.Name, "Propertie:", key, "Value:", value)
	}
}

func (i *wmiInstance) ExportDescription() {
	for key, _ := range i.Properties {
		fmt.Println("Propertie:", key)
	}
}

func getWmiClass(className string) (wmiClass, error) {
	var err error
	instances := map[string]interface{}{}
	c := wmiClass{Name: className}

	// init COM, oh yeah
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	SWbemLocator, _ := oleutil.CreateObject("WbemScripting.SWbemLocator")
	defer SWbemLocator.Release()

	wmi, _ := SWbemLocator.QueryInterface(ole.IID_IDispatch)
	defer wmi.Release()

	// service is a SWbemServices
	serviceRaw, _ := oleutil.CallMethod(wmi, "ConnectServer")
	service := serviceRaw.ToIDispatch()
	defer service.Release()

	// result is a SWBemObjectSet
	resultRaw, err := oleutil.CallMethod(service, "ExecQuery", "SELECT * FROM "+className)
	if err != nil {
		return c, err
	}

	result := resultRaw.ToIDispatch()
	defer result.Release()

	enumInstances, err := result.GetProperty("_NewEnum")
	enumInst, err := enumInstances.ToIUnknown().IEnumVARIANT(ole.IID_IEnumVariant)

	// Loop through all class instances
	for instancesRaw, length, err := enumInst.Next(1); length > 0; instancesRaw, length, err = enumInst.Next(1) {
		if err != nil {
			return c, err
		}
		instance := instancesRaw.ToIDispatch()
		instanceName, err := oleutil.GetProperty(instance, "Name")
		if err != nil {
			return c, err
		}
		props, err := instance.GetProperty("Properties_")
		if err != nil {
			return c, err
		}

		enumProperties, err := props.ToIDispatch().GetProperty("_NewEnum")
		if err != nil {
			return c, err
		}
		enumProps, err := enumProperties.ToIUnknown().IEnumVARIANT(ole.IID_IEnumVariant)
		if err != nil {
			return c, err
		}

		// Loop through all instance properties
		for propertiesRaw, length2, err := enumProps.Next(1); length2 > 0; propertiesRaw, length2, err = enumProps.Next(1) {
			if err != nil {
				return c, err
			}
			properties := propertiesRaw.ToIDispatch()
			propertieName, err := oleutil.GetProperty(properties, "Name")
			if err != nil {
				return c, err
			}
			propertieValue, _ := oleutil.GetProperty(instance, propertieName.ToString())
			if propertieValue.VT == 0x3 {
				instances[propertieName.ToString()] = propertieValue.Value()
				i := wmiInstance{Name: instanceName.ToString(), Properties: instances}
				c.Instances = append(c.Instances, i)
			}
		}
	}
	return c, nil
}

type CustomCollector struct {
	PropertyName *prometheus.Desc
}

// NewCustomCollector ...
func NewCustomCollector() (Collector, error) {
	const subsystem = "Custom"

	return &CustomCollector{
		PropertyName: prometheus.NewDesc(
			prometheus.BuildFQName(Namespace, subsystem, "PropertyName"),
			"PropertyName",
			nil,
			nil,
		),
	}, nil
}

// Collect sends the metric values for each metric
// to the provided prometheus Metric channel.
func (c *CustomCollector) Collect(ch chan<- prometheus.Metric) error {
	if desc, err := c.collect(ch); err != nil {
		log.Println("[ERROR] failed collecting Custom metrics:", desc, err)
		return err
	}
	return nil
}

func (c *CustomCollector) collect(ch chan<- prometheus.Metric) (*prometheus.Desc, error) {

	instances := map[string]interface{}{}
	wmiClassList := string(*customClassPath)
	wmiClassListArray := strings.Split(wmiClassList, ";")
	for _, wmiClassPath := range wmiClassListArray {
		wmiClassArray := strings.Split(wmiClassPath, ":")
		wmiNamespace := wmiClassArray[0]
		wmiClassName := wmiClassArray[1]
		wc := wmiClass{Name: wmiNamespace}

		// init COM
		ole.CoInitialize(0)
		defer ole.CoUninitialize()

		SWbemLocator, _ := oleutil.CreateObject("WbemScripting.SWbemLocator")
		defer SWbemLocator.Release()

		wmi, _ := SWbemLocator.QueryInterface(ole.IID_IDispatch)
		defer wmi.Release()

		// service is a SWbemServices
		serviceRaw, _ := oleutil.CallMethod(wmi, "ConnectServer", nil, wmiNamespace, nil, nil, "MS_409")
		service := serviceRaw.ToIDispatch()
		defer service.Release()

		// result is a SWBemObjectSet
		resultRaw, err := oleutil.CallMethod(service, "ExecQuery", "SELECT * FROM "+wmiClassName)
		if err != nil {
			return nil, err
		}

		result := resultRaw.ToIDispatch()
		defer result.Release()

		enumInstances, err := result.GetProperty("_NewEnum")
		enumInst, err := enumInstances.ToIUnknown().IEnumVARIANT(ole.IID_IEnumVariant)

		// Loop through all class instances
		for instancesRaw, length, err := enumInst.Next(1); length > 0; instancesRaw, length, err = enumInst.Next(1) {
			if err != nil {
				return nil, err
			}
			instance := instancesRaw.ToIDispatch()
			instancePath, err := oleutil.GetProperty(instance, "Path_")
			instanceRelativePath, err := instancePath.ToIDispatch().GetProperty("RelPath")
			if err != nil {
				return nil, err
			}
			instanceName := strings.Split(instanceRelativePath.ToString(), "\"")[1]

			props, err := instance.GetProperty("Properties_")
			if err != nil {
				return nil, err
			}

			enumProperties, err := props.ToIDispatch().GetProperty("_NewEnum")
			if err != nil {
				return nil, err
			}
			enumProps, err := enumProperties.ToIUnknown().IEnumVARIANT(ole.IID_IEnumVariant)
			if err != nil {
				return nil, err
			}

			// Loop through all instance properties
			for propertiesRaw, length2, err := enumProps.Next(1); length2 > 0; propertiesRaw, length2, err = enumProps.Next(1) {
				if err != nil {
					return nil, err
				}
				properties := propertiesRaw.ToIDispatch()
				propertieName, err := oleutil.GetProperty(properties, "Name")
				if err != nil {
					return nil, err
				}

				propertieValue, _ := oleutil.GetProperty(instance, propertieName.ToString())
				var value float64
				// Manage numerical values
				switch propertieValue.VT {
				case ole.VT_I1:
					value = float64(int8(propertieValue.Val))
				case ole.VT_UI1:
					value = float64(uint8(propertieValue.Val))
				case ole.VT_I2:
					value = float64(int16(propertieValue.Val))
				case ole.VT_UI2:
					value = float64(uint16(propertieValue.Val))
				case ole.VT_I4:
					value = float64(int32(propertieValue.Val))
				case ole.VT_UI4:
					value = float64(uint32(propertieValue.Val))
				case ole.VT_I8:
					value = float64(int64(propertieValue.Val))
				case ole.VT_UI8:
					value = float64(uint64(propertieValue.Val))
				case ole.VT_INT:
					value = float64(int(propertieValue.Val))
				case ole.VT_UINT:
					value = float64(uint(propertieValue.Val))
				case ole.VT_BOOL:
					if propertieValue.Val != 0 {
						value = float64(1)
					} else {
						value = float64(0)
					}
				default:
					continue
				}
				instances[propertieName.ToString()] = propertieValue.Value()
				i := wmiInstance{Name: instanceName, Properties: instances}
				wc.Instances = append(wc.Instances, i)
				name := strings.ToLower(propertieName.ToString())
				ch <- prometheus.MustNewConstMetric(
					prometheus.NewDesc(prometheus.BuildFQName(Namespace, strings.ToLower(wmiClassName), name), name, []string{"wmiinstance"}, nil),
					prometheus.CounterValue,
					value,
					instanceName,
				)
			}
		}
	}

	return nil, nil
}
