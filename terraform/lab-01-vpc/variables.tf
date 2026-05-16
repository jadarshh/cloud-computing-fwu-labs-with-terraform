variable "lab_name" {
  type    = string
  default = "lab-01-vpc"
}

variable "suffix" {
  type        = string
  description = "Lowercase identifier used in resource names (e.g. 'mukesh')"
  default     = "mukesh"
}

variable "display_name" {
  type        = string
  description = "Short display name used in tags (e.g. 'Mukesh')"
  default     = "Mukesh"
}

variable "full_name" {
  type        = string
  description = "Student full name used for lab documentation."
  default     = "Mukesh Pant"
}

variable "roll_number" {
  type        = string
  description = "Student roll number used for lab documentation."
  default     = "29"
}

variable "semester" {
  type        = string
  description = "Student semester used for lab documentation."
  default     = "VIII"
}

variable "region" {
  type        = string
  description = "AWS region where lab resources are created."
  default     = "ap-south-1"
}

variable "region_label" {
  type        = string
  description = "Human-readable AWS region name used for lab documentation."
  default     = "Asia Pacific (Mumbai)"
}

variable "institution" {
  type        = string
  description = "Institution name used for lab documentation."
  default     = "Far Western University"
}

variable "faculty" {
  type        = string
  description = "Faculty name used for lab documentation."
  default     = "Faculty of Engineering"
}

variable "subject_name" {
  type        = string
  description = "Subject name used for lab documentation."
  default     = "Cloud Computing"
}
