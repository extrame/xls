package xls

type XF interface {
	FormatNo() uint16
}

type Xf5 struct {
	Font      uint16
	Format    uint16
	Type      uint16
	Align     uint16
	Color     uint16
	Fill      uint16
	Border    uint16
	Linestyle uint16
}

func (x *Xf5) FormatNo() uint16 {
	return x.Format
}

type Xf8 struct {
	Font        uint16
	Format      uint16
	Type        uint16
	Align       byte
	Rotation    byte
	Ident       byte
	Usedattr    byte
	Linestyle   uint32
	Linecolor   uint32
	Groundcolor uint16
}

func (x *Xf8) FormatNo() uint16 {
	return x.Format
}
